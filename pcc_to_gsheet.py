import time
import os
import sys
import pandas as pd
import gspread
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
HEADLESS_MODE = True # 上傳 GitHub 時請設為 True

KEYWORDS = ["資源回收", "分選", "細分選場", "細分選廠", "細分類", "廢棄物"]
ORG_KEYWORDS = ["資源循環署", "環境管理署"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 
WORKSHEET_NAME = 'news'

# 基本查詢網址
TARGET_URL = "https://web.pcc.gov.tw/prkms/tender/common/basic/indexTenderBasic"

def init_driver():
    chrome_options = Options()
    if HEADLESS_MODE:
        chrome_options.add_argument("--headless") 
    
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        print(f"❌ 瀏覽器啟動失敗: {e}")
        sys.exit(1)

def search_pcc(driver, keyword, search_type):
    print(f"\n🔍 正在搜尋 [{('機關' if search_type=='org' else '標案')}]：{keyword} ...")
    
    try:
        driver.get(TARGET_URL)
        wait = WebDriverWait(driver, 20)

        # 1. 填入搜尋框
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
        
        # 3. 等待結果
        try:
            wait.until(EC.presence_of_element_located((By.ID, "tpam")))
            page_text = driver.find_element(By.TAG_NAME, "body").text
            if "無符合條件資料" in page_text or "無資料" in page_text:
                print(f"   -> 查無資料 (跳過)")
                return []
        except:
            print(f"   -> 載入超時或無表格 (跳過)")
            return []
        
        # 4. 抓取資料
        results = []
        rows = driver.find_elements(By.CSS_SELECTOR, "#tpam tbody tr")
        JUNK_TITLES = ["標案查詢", "決標查詢", "全文檢索", "公告日期查詢", "機關名稱查詢", "功能選項", "更正公告"]

        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 9: continue
                
            try:
                # [1] 機關名稱 (Org)
                org_name = cols[1].text.strip()
                
                # [6] 公告日期 (Date)
                date_str = cols[6].text.strip()
                
                # [7] 截止投標 (Deadline)
                deadline = cols[7].text.strip()
                
                # [8] 預算金額 (Budget)
                budget = cols[8].text.strip()
                
                # [2] 標案名稱 (Title) & 連結 (Link)
                tender_cell = cols[2]
                try:
                    link_elem = tender_cell.find_element(By.TAG_NAME, "a")
                    tender_link = link_elem.get_attribute("href")
                except:
                    tender_link = ""

                full_text = tender_cell.text 
                lines = full_text.split('\n')
                tender_name = max(lines, key=len).strip()

                if not tender_name or len(tender_name) < 2: continue
                if any(junk in tender_name for junk in JUNK_TITLES): continue

                results.append({
                    "Date": date_str,
                    "Org": org_name,        # 新欄位
                    "Title": tender_name,
                    "Link": tender_link,
                    "Deadline": deadline,   # 新欄位
                    "Budget": budget,       # 新欄位
                    "Tags": f"{('機關' if search_type=='org' else '標案')}-{keyword}",
                    "Source": "政府採購網"
                })
            except:
                continue 
        
        print(f"   -> 成功找到 {len(results)} 筆")
        return results

    except Exception as e:
        print(f"   ❌ 搜尋發生錯誤: {e}")
        return []

def upload_to_gsheet(df):
    """上傳至 Google Sheets"""
    print("\n☁️ 正在連線 Google Sheets...")
    
    if not os.path.exists(JSON_KEY_FILE):
        print(f"❌ 錯誤：找不到 key.json")
        return

    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
        client = gspread.authorize(creds)
        
        sheet = client.open_by_url(SHEET_URL).worksheet(WORKSHEET_NAME)
        existing_data = sheet.get_all_records()
        
        # 檢查防重複 (假設 Link 是第 4 欄，也就是 row['Link'])
        # 為了安全，我們檢查標題列，找出 'Link' 在第幾欄
        # 不過這裡我們先用簡單的 key 對應
        existing_links = set(str(row['Link']) for row in existing_data if 'Link' in row)
        
        new_rows = []
        for index, row in df.iterrows():
            if str(row['Link']) not in existing_links:
                # 欄位順序必須跟 Google Sheet 標題一模一樣
                # Date, Org, Title, Link, Deadline, Budget, Tags, Source
                row_data = [
                    row['Date'],
                    row['Org'],
                    row['Title'],
                    row['Link'],
                    row['Deadline'],
                    row['Budget'],
                    row['Tags'],
                    row['Source']
                ]
                new_rows.append(row_data)
                existing_links.add(str(row['Link']))
        
        if new_rows:
            sheet.append_rows(new_rows)
            print(f"✅ 成功上傳 {len(new_rows)} 筆新資料！")
        else:
            print("⚠️ 沒有新的不重複資料需上傳。")
            
    except Exception as e:
        print(f"❌ 上傳失敗: {e}")

def main():
    print("🚀 啟動爬蟲 (V21.0 獨立欄位版)...")
    driver = init_driver()
    all_data = []
    
    try:
        # 1. 搜尋機關
        print("\n--- 搜尋機關名稱 ---")
        for org in ORG_KEYWORDS:
            data = search_pcc(driver, org, search_type="org")
            all_data.extend(data)
            time.sleep(1)

        # 2. 搜尋標案
        print("\n--- 搜尋標案關鍵字 ---")
        for kw in KEYWORDS:
            data = search_pcc(driver, kw, search_type="name")
            all_data.extend(data)
            time.sleep(1)
            
    finally:
        print("\n🛑 關閉瀏覽器...")
        driver.quit() 
        
    if all_data:
        df = pd.DataFrame(all_data)
        df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
        print(f"\n📊 共抓取到 {len(df)} 筆有效資料")
        upload_to_gsheet(df)
    else:
        print("\n❌ 本次執行沒有找到任何標案。")

if __name__ == "__main__":
    main()
