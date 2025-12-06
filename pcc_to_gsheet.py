import time
import os
import sys
import pandas as pd
import gspread
import re
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

# 1. 搜尋清單
KEYWORDS = ["資源回收", "分選", "細分選場", "細分選廠", "細分類", "廢棄物"]
ORG_KEYWORDS = ["資源循環署", "環境管理署"]

# 2. Google Sheets 設定
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 
WORKSHEET_NAME = 'news'

# 3. 目標網址 (PIS 新版首頁 - 您指定要用的)
TARGET_URL = "https://web.pcc.gov.tw/pis/"

def init_driver():
    """初始化瀏覽器"""
    chrome_options = Options()
    # ⚠️ 雲端執行必開 headless
    chrome_options.add_argument("--headless") 
    
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    # 偽裝 User-Agent 避免被 PIS 擋
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        print(f"❌ 瀏覽器啟動失敗: {e}")
        sys.exit(1)

def clean_pis_title(raw_text):
    """
    PIS 標題清理專用函式
    PIS 的連結文字通常長這樣： "1130101(更正公告)\n環境部資源循環署...\n公開招標..."
    我們需要切掉前面的案號和後面的廢話，只留中間的標題。
    """
    if not raw_text: return ""
    
    # 1. 將換行符號取代為空格，方便處理
    text = raw_text.replace('\r', '').strip()
    lines = text.split('\n')
    
    # 2. 智慧挑選策略：
    # 通常 PIS 卡片連結有三行：案號、標題、狀態
    # 我們找出「最長」的那一行，通常就是標題
    best_line = max(lines, key=len)
    
    # 3. 如果找不到長句，就回傳原文字(去除換行)
    if len(best_line) < 4:
        return text.replace('\n', ' ')
        
    return best_line.strip()

def search_pis(driver, keyword, search_type):
    """
    PIS 通用搜尋邏輯 (使用單一搜尋框)
    """
    print(f"\n🔍 [PIS] 正在搜尋 ({search_type})：{keyword} ...")
    results = []
    
    try:
        driver.get(TARGET_URL)
        wait = WebDriverWait(driver, 20)

        # 1. 找到 PIS 首頁大搜尋框
        try:
            # 等待輸入框 (input type=text)
            input_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='text']")))
            
            # 使用 JavaScript 清空並點擊，確保游標在裡面
            driver.execute_script("arguments[0].click(); arguments[0].value = '';", input_box)
            
            # 輸入關鍵字並按 Enter (PIS 不一定有按鈕，按 Enter 最穩)
            input_box.send_keys(keyword)
            time.sleep(0.5)
            input_box.send_keys(Keys.ENTER)
            
        except Exception as e:
            print(f"   ⚠️ 找不到 PIS 搜尋框: {e}")
            return []

        # 2. 等待搜尋結果 (卡片)
        try:
            # 等待出現含有 'tender' 的連結
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[href*='tender']")))
            time.sleep(3) # 等待資料渲染
        except:
            print(f"   -> 查無資料 (或載入超時)")
            return []
        
        # 3. 抓取資料
        # 抓取所有包含 tender 的連結 (這是 PIS 標案卡的特徵)
        links_elements = driver.find_elements(By.CSS_SELECTOR, "a[href*='tender']")
        
        # PIS 列表頁比較難抓日期，我們暫用今日日期，或者嘗試從文字中提取
        date_str = datetime.now().strftime("%Y-%m-%d")

        print(f"   -> 偵測到 {len(links_elements)} 個項目，開始過濾...")

        for elem in links_elements:
            try:
                # 抓取原始文字
                raw_text = elem.get_attribute("innerText")
                link = elem.get_attribute("href")
                
                # 清洗標題 (使用上面的專用函式)
                title = clean_pis_title(raw_text)
                
                # 過濾無效資料
                if len(title) < 4: continue
                # 過濾系統連結
                if "更多" in title or "機關" in title: continue

                # 去重
                if not any(d['Link'] == link for d in results):
                    results.append({
                        "Date": date_str,
                        "Title": title,
                        "Link": link,
                        "Tags": f"PIS-{search_type}-{keyword}",
                        "Source": "政府採購網PIS"
                    })
                
                # 每個關鍵字只抓前 15 筆
                if len(results) >= 15: break
            except:
                continue
        
        print(f"   -> 成功提取 {len(results)} 筆有效資料")
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
    print("🚀 啟動 PIS 爬蟲 (V12.0 標題修復版)...")
    driver = init_driver()
    all_data = []
    
    try:
        # 1. 搜尋機關名稱
        print("\n--- 開始搜尋機關 ---")
        for org in ORG_KEYWORDS:
            data = search_pis(driver, org, search_type="機關")
            all_data.extend(data)
            time.sleep(2)

        # 2. 搜尋標案關鍵字
        print("\n--- 開始搜尋標案關鍵字 ---")
        for kw in KEYWORDS:
            data = search_pis(driver, kw, search_type="標案")
            all_data.extend(data)
            time.sleep(2)
            
    finally:
        print("\n🛑 關閉瀏覽器...")
        driver.quit()
        
    if all_data:
        df = pd.DataFrame(all_data)
        # 根據網址去重
        df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
        
        print(f"\n📊 共抓取到 {len(df)} 筆資料，準備上傳...")
        upload_to_gsheet(df)
    else:
        print("\n❌ 本次執行沒有找到任何標案。")

if __name__ == "__main__":
    main()
