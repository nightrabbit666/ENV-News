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

# 1. Google Sheets 設定
# 程式會自動抓取當前目錄下的 key.json (由 GitHub Actions 產生)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')

# 2. 試算表設定 (請確認您的網址與工作表名稱)
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 
WORKSHEET_NAME = 'news'

# 3. 搜尋關鍵字
KEYWORDS = ["資源回收", "分選", "細分選場", "細分選廠", "細分類", "廢棄物"]
ORG_KEYWORDS = ["資源循環署", "環境管理署"]
TARGET_URL = "https://web.pcc.gov.tw/pis/"

def init_driver():
    """初始化瀏覽器 (雲端專用設定)"""
    chrome_options = Options()
    # ⚠️ 強制開啟無頭模式 (雲端環境必備)
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

def search_pis(driver, keyword):
    """PIS 搜尋邏輯"""
    print(f"\n🔍 [PIS] 正在搜尋：{keyword} ...")
    results = []
    
    try:
        driver.get(TARGET_URL)
        wait = WebDriverWait(driver, 20)

        # 1. 搜尋
        try:
            input_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='text']")))
            input_box.clear()
            input_box.send_keys(keyword)
            time.sleep(0.5)
            input_box.send_keys(Keys.ENTER)
        except Exception as e:
            print(f"   ⚠️ 找不到搜尋框: {e}")
            return []

        # 2. 等待結果
        try:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "a")))
            time.sleep(2)
        except:
            print(f"   -> 查無資料或載入超時")
            return []
        
        # 3. 抓取資料
        links_elements = driver.find_elements(By.CSS_SELECTOR, "a[href*='tender']")
        date_str = datetime.now().strftime("%Y-%m-%d")

        for elem in links_elements:
            try:
                title = elem.text.strip()
                link = elem.get_attribute("href")
                if len(title) < 4: continue
                
                # 簡單去重檢查
                if not any(d['Link'] == link for d in results):
                    results.append({
                        "Date": date_str,
                        "Tags": f"PIS搜尋-{keyword}",
                        "Title": title,
                        "Link": link,
                        "Source": "政府採購網PIS"
                    })
                if len(results) >= 10: break
            except: continue
            
        print(f"   -> 成功提取 {len(results)} 筆有效資料")
        return results

    except Exception as e:
        print(f"   ❌ 搜尋發生錯誤: {e}")
        return []

def upload_to_gsheet(df):
    """上傳至 Google Sheets"""
    print("\n☁️ 正在連線 Google Sheets...")
    
    if not os.path.exists(JSON_KEY_FILE):
        print(f"❌ 錯誤：找不到 key.json！(請確認 GitHub Secrets 是否設定正確)")
        return

    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
        client = gspread.authorize(creds)
        
        sheet = client.open_by_url(SHEET_URL).worksheet(WORKSHEET_NAME)
        existing_data = sheet.get_all_records()
        # 建立現有連結的集合，用於防重複
        existing_links = set(str(row['Link']) for row in existing_data if 'Link' in row)
        
        new_rows = []
        for index, row in df.iterrows():
            if str(row['Link']) not in existing_links:
                # 欄位順序: Date, Tags, Title, Link, Source
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
    print("🚀 啟動 PIS 雲端爬蟲...")
    driver = init_driver()
    all_data = []
    
    try:
        search_list = KEYWORDS + ORG_KEYWORDS
        for kw in search_list:
            data = search_pis(driver, kw)
            all_data.extend(data)
            time.sleep(2)
    finally:
        print("🛑 關閉瀏覽器...")
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
