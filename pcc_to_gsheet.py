import time
import os
import sys
import pandas as pd
import gspread
import traceback
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- 設定區 ---
KEYWORDS = ["資源回收", "分選", "細分選場", "細分選廠", "細分類", "廢棄物"]
ORG_KEYWORDS = ["資源循環署", "環境管理署"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 
WORKSHEET_NAME = 'news'
LOG_SHEET_NAME = 'logs' # 新增：日誌工作表名稱

TARGET_URL = "https://web.pcc.gov.tw/prkms/tender/common/basic/indexTenderBasic"

def get_google_client():
    if not os.path.exists(JSON_KEY_FILE):
        raise FileNotFoundError(f"找不到 key.json")
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    return gspread.authorize(creds)

def log_to_sheet(status, message):
    """寫入系統日誌"""
    print(f"[{status}] {message}")
    try:
        client = get_google_client()
        sheet = client.open_by_url(SHEET_URL).worksheet(LOG_SHEET_NAME)
        # 寫入: 時間, 狀態, 訊息
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sheet.append_row([timestamp, status, message])
    except Exception as e:
        print(f"❌ 無法寫入日誌: {e}")

def init_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
    
    try:
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        raise Exception(f"瀏覽器啟動失敗: {e}")

def search_pcc(driver, keyword, search_type):
    print(f"\n🔍 搜尋 [{search_type}]：{keyword}")
    try:
        driver.get(TARGET_URL)
        wait = WebDriverWait(driver, 20)

        if search_type == "name":
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "tenderName")))
            driver.find_element(By.NAME, "orgName").clear()
        else:
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "orgName")))
            driver.find_element(By.NAME, "tenderName").clear()
            
        input_box.clear()
        input_box.send_keys(keyword)
        time.sleep(0.5) 
        input_box.send_keys(Keys.ENTER)
        
        try:
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tb_01")))
            page_text = driver.find_element(By.TAG_NAME, "body").text
            if "無符合條件資料" in page_text or "無資料" in page_text:
                return []
        except:
            return []
        
        results = []
        rows = driver.find_elements(By.CSS_SELECTOR, ".tb_01 tbody tr")
        JUNK_TITLES = ["標案查詢", "決標查詢", "全文檢索", "公告日期查詢", "機關名稱查詢", "功能選項", "更正公告"]

        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 7: continue
            try:
                org_name = cols[1].text.strip()
                date_str = cols[6].text.strip()
                
                links_in_cell = cols[2].find_elements(By.TAG_NAME, "a")
                if links_in_cell:
                    longest_link = max(links_in_cell, key=lambda x: len(x.text.strip()))
                    tender_name = longest_link.text.strip()
                    tender_link = longest_link.get_attribute("href")
                else:
                    tender_name = cols[2].text.strip()
                    tender_link = ""

                if not tender_name or len(tender_name) < 2: continue
                if any(junk in tender_name for junk in JUNK_TITLES): continue

                results.append({
                    "Date": date_str,
                    "Org": org_name,
                    "Title": tender_name,
                    "Link": tender_link,
                    "Deadline": cols[7].text.strip() if len(cols) > 7 else "",
                    "Budget": cols[8].text.strip() if len(cols) > 8 else "",
                    "Tags": f"{('機關' if search_type=='org' else '標案')}-{keyword}",
                    "Source": "政府採購網"
                })
            except:
                continue 
        return results
    except Exception as e:
        print(f"   ❌ 搜尋單項錯誤: {e}")
        return []

def upload_to_gsheet(df):
    print("\n☁️ 上傳 Google Sheets...")
    client = get_google_client()
    sheet = client.open_by_url(SHEET_URL).worksheet(WORKSHEET_NAME)
    existing_data = sheet.get_all_records()
    existing_links = set(str(row['Link']) for row in existing_data if 'Link' in row)
    
    new_rows = []
    for index, row in df.iterrows():
        if str(row['Link']) not in existing_links:
            row_data = [
                row['Date'], row['Org'], row['Title'], row['Link'],
                row['Deadline'], row['Budget'], row['Tags'], row['Source']
            ]
            new_rows.append(row_data)
            existing_links.add(str(row['Link']))
    
    if new_rows:
        sheet.append_rows(new_rows)
        return len(new_rows)
    return 0

def main():
    print("🚀 啟動爬蟲 (V23.0 錯誤回報版)...")
    try:
        driver = init_driver()
        all_data = []
        
        for org in ORG_KEYWORDS:
            all_data.extend(search_pcc(driver, org, "org"))
            time.sleep(1)

        for kw in KEYWORDS:
            all_data.extend(search_pcc(driver, kw, "name"))
            time.sleep(1)
            
        driver.quit()
        
        msg = "今日無新資料"
        if all_data:
            df = pd.DataFrame(all_data)
            df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
            count = upload_to_gsheet(df)
            msg = f"成功執行，新增 {count} 筆資料 (共抓取 {len(df)} 筆)"
        
        # ✅ 成功：寫入 Success 日誌
        log_to_sheet("SUCCESS", msg)

    except Exception as e:
        # ❌ 失敗：寫入 Error 日誌 (包含詳細錯誤原因)
        error_msg = f"程式崩潰: {str(e)}\n{traceback.format_exc()}"
        log_to_sheet("ERROR", error_msg)
        sys.exit(1)

if __name__ == "__main__":
    main()
