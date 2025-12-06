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
# 標案名稱關鍵字
KEYWORDS = ["資源回收", "分選", "細分選場", "細分選廠", "細分類", "廢棄物"]
# 機關名稱關鍵字
ORG_KEYWORDS = ["資源循環署", "環境管理署"]

# 2. Google Sheets 設定
# 自動抓取當前目錄下的 key.json
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')

# ⚠️ 請確認您的試算表網址與工作表名稱
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 
WORKSHEET_NAME = 'news'

# 使用「基本查詢」網址 (因為只有這裡可以區分機關與標案名稱)
TARGET_URL = "https://web.pcc.gov.tw/prkms/tender/common/basic/"

def init_driver():
    """初始化瀏覽器"""
    chrome_options = Options()
    # ⚠️ 強制開啟無頭模式 (GitHub Actions 必備)
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
    search_type: "name" (標案名稱) / "org" (機關名稱)
    """
    print(f"\n🔍 正在搜尋 [{('機關' if search_type=='org' else '標案')}]：{keyword} ...")
    
    try:
        driver.get(TARGET_URL)
        wait = WebDriverWait(driver, 20)

        # 1. 根據類型找到對應輸入框
        if search_type == "name":
            # 找「標案名稱」輸入框
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "tenderName")))
        else:
            # 找「機關名稱」輸入框
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "orgName")))
            
        input_box.clear()
        input_box.send_keys(keyword)
        
        # 2. 點擊「查詢」按鈕 (確保點到最下面那個，而非旁邊的小幫手)
        search_btn = driver.find_element(By.CSS_SELECTOR, "div.buttons input[name='search']")
        driver.execute_script("arguments[0].click();", search_btn)
        
        # 3. 等待結果
        try:
            # 等待表格出現 (最多等 5 秒)
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CLASS_NAME, "tb_01")))
        except:
            print(f"   -> 查無資料")
            return []
        
        # 4. 抓取資料 (精準定位欄位)
        results = []
        rows = driver.find_elements(By.CSS_SELECTOR, ".tb_01 tbody tr")
        
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            # 欄位檢查：基本查詢通常有 9 個欄位
            # [1]機關名稱, [2]標案案號, [3]標案名稱(含連結)... [6]公告日期
            if len(cols) < 8: continue
                
            try:
                # --- 關鍵修正：欄位抓取 ---
                # 抓取第 2 欄 (Index 1): 機關名稱
                org_name = cols[1].text.strip()
                
                # 抓取第 4 欄 (Index 3): 標案名稱與連結
                tender_link_elem = cols[3].find_element(By.TAG_NAME, "a")
                tender_name = tender_link_elem.text.strip()
                tender_link = tender_link_elem.get_attribute("href")
                
                # 抓取第 7 欄 (Index 6): 公告日期
                date_str = cols[6].text.strip()
                
                # 簡單過濾：只抓今年的，避免抓到陳年舊案 (可選)
                # if "114" not in date_str: continue

                results.append({
                    "Date": date_str,
                    "Title": tender_name,  # 這裡確保抓到的是標案名稱
                    "Link": tender_link,
                    "Tags": f"{('機關' if search_type=='org' else '關鍵字')}-{keyword}",
                    "Source": org_name     # 來源欄位填入機關名稱
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
    print("🚀 啟動政府採購網爬蟲 (V5.0 精準版)...")
    driver = init_driver()
    all_data = []
    
    try:
        # 1. 第一輪：搜尋機關名稱
        print("\n--- 開始搜尋機關 ---")
        for org in ORG_KEYWORDS:
            data = search_pcc(driver, org, search_type="org")
            all_data.extend(data)
            time.sleep(1)

        # 2. 第二輪：搜尋標案關鍵字
        print("\n--- 開始搜尋標案關鍵字 ---")
        for kw in KEYWORDS:
            data = search_pcc(driver, kw, search_type="name")
            all_data.extend(data)
            time.sleep(1)
            
    finally:
        print("\n🛑 關閉瀏覽器...")
        driver.quit()
        
    if all_data:
        df = pd.DataFrame(all_data)
        # 根據網址去重 (避免機關跟關鍵字搜到同一個)
        df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
        
        print(f"\n📊 共抓取到 {len(df)} 筆資料，準備上傳...")
        upload_to_gsheet(df)
    else:
        print("\n❌ 本次執行沒有找到任何標案。")

if __name__ == "__main__":
    main()
