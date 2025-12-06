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
# 自動抓取當前目錄下的 key.json (由 GitHub Actions 產生)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')

# ⚠️ 請確認您的試算表網址與工作表名稱
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 
WORKSHEET_NAME = 'news'

# 3. 目標網址 (基本查詢首頁 - 您指定的正確網址)
TARGET_URL = "https://web.pcc.gov.tw/prkms/tender/common/basic/indexTenderBasic"

def init_driver():
    """初始化瀏覽器"""
    chrome_options = Options()
    
    # ⚠️【關鍵修正】強制開啟無頭模式
    # 這行絕對不能被註解掉，否則 GitHub Actions 會直接報錯 (NoneType error)
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
        # 如果瀏覽器沒啟動成功，直接結束程式，避免後續報錯
        sys.exit(1)

def search_pcc(driver, keyword, search_type):
    """
    執行搜尋 (基本查詢介面專用)
    search_type: "name" (標案名稱) / "org" (機關名稱)
    """
    print(f"\n🔍 正在搜尋 [{('機關' if search_type=='org' else '標案')}]：{keyword} ...")
    
    try:
        driver.get(TARGET_URL)
        wait = WebDriverWait(driver, 20)

        # 1. 鎖定輸入框 (互斥邏輯：填一個，清空另一個)
        if search_type == "name":
            # 填入 @標案名稱
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "tenderName")))
            # 清空機關名稱
            driver.find_element(By.NAME, "orgName").clear()
        else:
            # 填入 @機關名稱
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "orgName")))
            # 清空標案名稱
            driver.find_element(By.NAME, "tenderName").clear()
            
        input_box.clear()
        input_box.send_keys(keyword)
        
        # 2. 點擊「查詢」按鈕 (紅框處)
        # 鎖定 form 裡面的查詢按鈕，避免點到旁邊的小幫手
        search_btn = driver.find_element(By.CSS_SELECTOR, "input[name='search']")
        driver.execute_script("arguments[0].click();", search_btn)
        
        # 3. 等待結果 & 嚴格過濾
        try:
            # 等待表格出現
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tb_01")))
            
            # ★ 關鍵檢查：如果網頁顯示「無符合條件資料」，直接跳過！
            # 這樣就絕對不會去抓到上面的選單按鈕
            page_text = driver.find_element(By.TAG_NAME, "body").text
            if "無符合條件資料" in page_text or "無資料" in page_text:
                print(f"   -> 查無資料 (跳過)")
                return []
        except:
            print(f"   -> 載入超時或無表格 (跳過)")
            return []
        
        # 4. 抓取資料 (精準抓取表格內容)
        results = []
        rows = driver.find_elements(By.CSS_SELECTOR, ".tb_01 tbody tr")
        
        # 垃圾標題黑名單 (雙重保險)
        JUNK_TITLES = ["標案查詢", "決標查詢", "全文檢索", "公告日期查詢", "機關名稱查詢"]

        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            # 欄位數量檢查：基本查詢表格通常有 9 欄
            if len(cols) < 7: continue
                
            try:
                # [1] 機關名稱
                org_name = cols[1].text.strip()
                
                # [6] 公告日期
                date_str = cols[6].text.strip()
                
                # [2] 標案名稱與連結
                # 這裡同時有「案號」跟「名稱」，我們用長度判斷抓出名稱
                links_in_cell = cols[2].find_elements(By.TAG_NAME, "a")
                
                tender_name = ""
                tender_link = ""
                
                if links_in_cell:
                    # 找出文字最長的連結 (排除短案號)
                    longest_link = max(links_in_cell, key=lambda x: len(x.text.strip()))
                    tender_name = longest_link.text.strip()
                    tender_link = longest_link.get_attribute("href")
                else:
                    # 沒連結就抓純文字
                    tender_name = cols[2].text.strip()

                # ★ 防垃圾過濾器
                # 1. 如果標題太短或空的 -> 丟掉
                if not tender_name or len(tender_name) < 2: continue
                # 2. 如果標題是選單文字 -> 丟掉
                if any(junk in tender_name for junk in JUNK_TITLES): continue

                results.append({
                    "Date": date_str,
                    "Title": tender_name,
                    "Link": tender_link,
                    "Tags": f"{('機關' if search_type=='org' else '標案')}-{keyword}",
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
    print("🚀 啟動政府採購網爬蟲 (V15.0 雲端修正版)...")
    driver = init_driver()
    all_data = []
    
    try:
        # 1. 搜尋機關
        print("\n--- 搜尋機關名稱 ---")
        for org in ORG_KEYWORDS:
            data = search_pcc(driver, org, search_type="org")
            all_data.extend(data)
            time.sleep(2)

        # 2. 搜尋標案
        print("\n--- 搜尋標案關鍵字 ---")
        for kw in KEYWORDS:
            data = search_pcc(driver, kw, search_type="name")
            all_data.extend(data)
            time.sleep(2)
            
    finally:
        print("\n🛑 關閉瀏覽器...")
        if driver:
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
