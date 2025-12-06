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
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- 設定區 ---

# 1. 搜尋清單
KEYWORDS = ["資源回收", "分選", "細分選場", "細分選廠", "細分類", "廢棄物"]
ORG_KEYWORDS = ["資源循環署", "環境管理署"]

# 2. Google Sheets 設定
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 
WORKSHEET_NAME = 'news'

# 3. 目標網址
# 雖然您是在 PIS 系統操作，但該「找標案」介面的真實表格位址是這裡
# 直接讓機器人連這裡，可以保證找到對應的輸入框，且畫面與您的截圖一致
TARGET_URL = "https://web.pcc.gov.tw/prkms/tender/common/basic/"

def init_driver():
    """初始化瀏覽器"""
    chrome_options = Options()
    # ⚠️ 雲端執行 (GitHub Actions) 必開無頭模式
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
    """
    執行搜尋
    search_type: "name" (搜標案名稱) / "org" (搜機關名稱)
    """
    print(f"\n🔍 正在搜尋 [{('機關' if search_type=='org' else '標案')}]：{keyword} ...")
    
    try:
        driver.get(TARGET_URL)
        wait = WebDriverWait(driver, 20)

        # 1. 根據搜尋類型，填入正確的格子 (對應您的截圖)
        if search_type == "name":
            # 填入 @標案名稱 (tenderName)
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "tenderName")))
            # 確保機關名稱是空的，以免干擾
            driver.find_element(By.NAME, "orgName").clear()
        else:
            # 填入 @機關名稱 (orgName)
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "orgName")))
            # 確保標案名稱是空的
            driver.find_element(By.NAME, "tenderName").clear()
            
        input_box.clear()
        input_box.send_keys(keyword)
        
        # 2. 點擊下方的「查詢」按鈕 (紅框處)
        # 我們利用 CSS Selector 精準定位那個位於 buttons 區塊內的查詢按鈕
        search_btn = driver.find_element(By.CSS_SELECTOR, "div.buttons input[name='search']")
        driver.execute_script("arguments[0].click();", search_btn)
        
        # 3. 等待結果表格出現
        try:
            # 等待表格 (class="tb_01")
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CLASS_NAME, "tb_01")))
        except:
            print(f"   -> 查無資料")
            return []
        
        # 4. 抓取資料 (根據您的截圖 image_446cb3.png 校正欄位)
        results = []
        rows = driver.find_elements(By.CSS_SELECTOR, ".tb_01 tbody tr")
        
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            # 欄位數量檢查
            if len(cols) < 7: continue
                
            try:
                # [1] 機關名稱
                org_name = cols[1].text.strip()
                
                # [2] 標案案號 / 標案名稱 (這格裡面有連結 <a>)
                tender_link_elem = cols[2].find_element(By.TAG_NAME, "a")
                tender_name = tender_link_elem.text.strip()
                tender_link = tender_link_elem.get_attribute("href")
                
                # [6] 公告日期
                date_str = cols[6].text.strip()
                
                # 排除空資料
                if not tender_name: continue

                results.append({
                    "Date": date_str,
                    "Title": tender_name,
                    "Link": tender_link,
                    "Tags": f"{('機關' if search_type=='org' else '標案關鍵字')}-{keyword}",
                    "Source": org_name
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
        print(f"❌ 錯誤：找不到 key.json！路徑: {JSON_KEY_FILE}")
        return

    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
        client = gspread.authorize(creds)
        
        sheet = client.open_by_url(SHEET_URL).worksheet(WORKSHEET_NAME)
        existing_data = sheet.get_all_records()
        existing_links = set(str(row['Link']) for row in existing_data if 'Link' in row)
        
        new_rows = []
        for index, row in df.iterrows():
            if str(row['Link']) not in existing_links:
                # 欄位: Date, Tags, Title, Link, Source
                row_data = [row['Date'], row['Tags'], row['Title'], row['Link'], row['Source']]
                new_rows.append(row_data)
                existing_links.add(str(row['Link']))
        
        if new_rows:
            sheet.append_rows(new_rows)
            print(f"✅ 成功上傳 {len(new_rows)} 筆新資料到雲端！")
        else:
            print("⚠️ 沒有新的不重複資料需上傳。")
            
    except Exception as e:
        print(f"❌ 上傳 Google Sheets 失敗: {e}")

def main():
    print("🚀 啟動爬蟲 (V7.0 截圖對應版)...")
    driver = init_driver()
    all_data = []
    
    try:
        # 1. 搜機關 (填入 機關名稱 框)
        print("\n--- 搜尋機關名稱 ---")
        for org in ORG_KEYWORDS:
            data = search_pcc(driver, org, search_type="org")
            all_data.extend(data)
            time.sleep(1)

        # 2. 搜標案 (填入 標案名稱 框)
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
        print(f"\n📊 共抓取到 {len(df)} 筆資料，準備上傳...")
        upload_to_gsheet(df)
    else:
        print("\n❌ 本次執行沒有找到任何標案。")

if __name__ == "__main__":
    main()
