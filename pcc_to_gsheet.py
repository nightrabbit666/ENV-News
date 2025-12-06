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

# 搜尋清單 (PIS 搜尋框通用，所以我們會依序丟進去搜)
KEYWORDS = ["資源回收", "分選", "細分選場", "細分選廠", "細分類", "廢棄物"]
ORG_KEYWORDS = ["資源循環署", "環境管理署"]

# Google Sheets 設定
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 
WORKSHEET_NAME = 'news'

# PIS 首頁
TARGET_URL = "https://web.pcc.gov.tw/pis/"

def init_driver():
    chrome_options = Options()
    # ⚠️ 雲端必開 headless
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    # 偽裝成一般使用者
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        print(f"❌ 瀏覽器啟動失敗: {e}")
        sys.exit(1)

def search_pis(driver, keyword, search_type):
    print(f"\n🔍 [PIS] 正在搜尋 ({search_type})：{keyword} ...")
    results = []
    
    try:
        driver.get(TARGET_URL)
        wait = WebDriverWait(driver, 20)

        # 1. 找到搜尋框 (PIS 首頁正中間那個)
        # 通常是 input type="text" 且 placeholder 包含 "關鍵字"
        try:
            # 等待輸入框出現
            input_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='text']")))
            
            # 確保清空並輸入
            input_box.click()
            input_box.clear()
            input_box.send_keys(keyword)
            time.sleep(0.5)
            # PIS 按 Enter 最快，不用找按鈕
            input_box.send_keys(Keys.ENTER)
            
        except Exception as e:
            print(f"   ⚠️ 找不到 PIS 搜尋框: {e}")
            return []

        # 2. 等待搜尋結果 (動態載入)
        # 我們等待頁面上出現標案連結 (通常 href 包含 tender)
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[href*='tender']")))
            time.sleep(3) # PIS 是動態的，多等一下讓資料跑出來
        except:
            print(f"   -> 查無資料 (或載入過久)")
            return []
        
        # 3. 抓取資料 (針對 PIS 卡片介面)
        # PIS 的結果通常是一張張卡片，標題是 <a> 標籤
        links_elements = driver.find_elements(By.CSS_SELECTOR, "a[href*='tender']")
        
        # PIS 列表頁通常不會直接顯示日期，我們用「今天」作為抓取日期
        # 或者嘗試抓取連結旁邊的文字
        date_str = datetime.now().strftime("%Y-%m-%d")

        for elem in links_elements:
            try:
                title = elem.text.strip()
                link = elem.get_attribute("href")
                
                # 過濾掉太短的文字 (例如 "更多...")
                if len(title) < 5: continue
                
                # PIS 有時候會搜出不相關的，簡單過濾
                # 如果是搜標案，標題最好包含關鍵字；如果是搜機關，則放寬
                
                if not any(d['Link'] == link for d in results):
                    results.append({
                        "Date": date_str,
                        "Title": title,
                        "Link": link,
                        "Tags": f"PIS-{search_type}-{keyword}",
                        "Source": "政府採購網PIS"
                    })
                
                # 每個關鍵字只抓前 15 筆，避免超時
                if len(results) >= 15: break
            except:
                continue
        
        print(f"   -> 成功提取 {len(results)} 筆有效資料")
        return results

    except Exception as e:
        print(f"   ❌ 搜尋發生錯誤: {e}")
        return []

def upload_to_gsheet(df):
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
        existing_links = set(str(row['Link']) for row in existing_data if 'Link' in row)
        
        new_rows = []
        for index, row in df.iterrows():
            if str(row['Link']) not in existing_links:
                row_data = [row['Date'], row['Tags'], row['Title'], row['Link'], row['Source']]
                new_rows.append(row_data)
                existing_links.add(str(row['Link']))
        
        if new_rows:
            sheet.append_rows(new_rows)
            print(f"✅ 成功上傳 {len(new_rows)} 筆新資料！")
        else:
            print("⚠️ 沒有新的不重複資料需上傳。")
            
    except Exception as e:
        print(f"❌ 上傳 Google Sheets 失敗: {e}")

def main():
    print("🚀 啟動 PIS 暴力搜尋爬蟲 (V8.0)...")
    driver = init_driver()
    all_data = []
    
    try:
        # 1. 搜尋機關
        # PIS 搜尋框很聰明，輸入機關名稱也能搜到該機關的標案
        for org in ORG_KEYWORDS:
            data = search_pis(driver, org, search_type="機關")
            all_data.extend(data)
            time.sleep(2)

        # 2. 搜尋標案關鍵字
        for kw in KEYWORDS:
            data = search_pis(driver, kw, search_type="標案")
            all_data.extend(data)
            time.sleep(2)
            
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
