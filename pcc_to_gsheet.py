import time
import os
import sys
import pandas as pd
import gspread
import traceback
import requests
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- 【切換開關】 (V23.1 功能回歸) ---
# ★ False = 本機看畫面 (除錯用)
# ★ True  = 雲端背景執行 (上傳 GitHub 前請改回 True)
HEADLESS_MODE = True
    
# --- 設定區 ---
# Google Chat 設定
# 如果是在本機測試，引號內直接貼上網址
# 如果是上傳 GitHub，建議寫 os.environ.get('GOOGLE_CHAT_WEBHOOK')
GOOGLE_CHAT_WEBHOOK = "https://chat.googleapis.com/v1/spaces/AAQAbfa7gJQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=N4OegGZLJ2y1ANxt41jIFf57RaGV4TI3Vw_GyHzdzeU"

# 預設關鍵字
DEFAULT_KEYWORDS = ["資源回收", "分選", "細分選場", "細分選廠", "細分類", "廢棄物"]
DEFAULT_ORG_KEYWORDS = ["資源循環署", "環境管理署"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 

WORKSHEET_NAME = 'news'
LOG_SHEET_NAME = 'logs'
CONFIG_SHEET_NAME = 'Config'
HISTORY_SHEET_NAME = 'history'

# 目標網址
TARGET_URL = "https://web.pcc.gov.tw/prkms/tender/common/basic/indexTenderBasic"

# --- 基礎建設函式 ---

def get_google_client():
    if not os.path.exists(JSON_KEY_FILE):
        raise FileNotFoundError(f"找不到 key.json")
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    return gspread.authorize(creds)

def log_to_sheet(status, message):
    """寫入系統日誌 (V27 功能)"""
    print(f"[{status}] {message}")
    try:
        client = get_google_client()
        sheet = client.open_by_url(SHEET_URL).worksheet(LOG_SHEET_NAME)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sheet.append_row([timestamp, status, message])
    except Exception as e:
        print(f"❌ 日誌寫入失敗: {e}")

def send_alert(message):
   def send_google_chat(new_data_count, df_new):
    """發送 Google Chat 通知 (V31.0 新增)"""
    if not GOOGLE_CHAT_WEBHOOK: return

    print("📲 準備發送 Google Chat 通知...")
    today = datetime.now().strftime("%Y/%m/%d")
    
    # 1. 標題
    text = f"🔔 *【標案戰情快訊】 {today}*\n"
    text += f"發現 {new_data_count} 筆新資料：\n"
    text += "━━━━━━━━━━━━━━\n"

    # 2. 列表內容 (只列出前 10 筆以免訊息太長)
    count = 0
    for index, row in df_new.iterrows():
        count += 1
        if count > 10:
            text += f"\n...(還有 {new_data_count - 10} 筆，請至儀表板查看)"
            break
        
        # 簡單排版：[機關] 標題
        title = row['Title'][:30] + "..." if len(row['Title']) > 30 else row['Title']
        text += f"{count}. [{row['Org']}] {title}\n"
        text += f"   💰 {row['Budget']} | ⏳ {row['Deadline']}\n"
        text += f"   🔗 <{row['Link']}|點擊查看>\n\n"

    # 3. 結尾
    text += "━━━━━━━━━━━━━━\n"
    # 這裡記得換成您的儀表板網址
    text += f"📊 <https://nightrabbit666.github.io/ENV-News/index.html|查看完整戰情儀表板>"

    # 4. 發送請求
    try:
        response = requests.post(
            GOOGLE_CHAT_WEBHOOK, 
            json={"text": text}
        )
        if response.status_code == 200:
            print("✅ Google Chat 發送成功！")
        else:
            print(f"❌ Google Chat 發送失敗: {response.text}")
    except Exception as e:
        print(f"❌ Google Chat 連線錯誤: {e}")
    """讀取雲端關鍵字 (V26 功能)"""
    try:
        client = get_google_client()
        sheet = client.open_by_url(SHEET_URL).worksheet(CONFIG_SHEET_NAME)
        records = sheet.get_all_records()
        keywords = [r['Keyword'] for r in records if r['Type'] == '標案' and r['Keyword']]
        orgs = [r['Keyword'] for r in records if r['Type'] == '機關' and r['Keyword']]
        
        if not keywords: keywords = DEFAULT_KEYWORDS
        if not orgs: orgs = DEFAULT_ORG_KEYWORDS
        return keywords, orgs
    except:
        return DEFAULT_KEYWORDS, DEFAULT_ORG_KEYWORDS

def archive_old_records():
    """自動封存舊資料 (V26 功能)"""
    print("\n📦 檢查資料封存...")
    try:
        client = get_google_client()
        doc = client.open_by_url(SHEET_URL)
        news_sheet = doc.worksheet(WORKSHEET_NAME)
        try:
            history_sheet = doc.worksheet(HISTORY_SHEET_NAME)
        except:
            return 

        all_records = news_sheet.get_all_records()
        if not all_records: return

        deadline = datetime.now() - timedelta(days=180)
        rows_keep, rows_archive = [], []
        header = news_sheet.row_values(1)

        for row in all_records:
            try:
                d_str = str(row['Date'])
                if '/' in d_str:
                    parts = d_str.split('/')
                    r_date = datetime(int(parts[0]) + 1911, int(parts[1]), int(parts[2]))
                    if r_date < deadline:
                        rows_archive.append(list(row.values()))
                    else:
                        rows_keep.append(list(row.values()))
                else:
                    rows_keep.append(list(row.values()))
            except:
                rows_keep.append(list(row.values()))

        if rows_archive:
            history_sheet.append_rows(rows_archive)
            news_sheet.clear()
            news_sheet.append_row(header)
            if rows_keep:
                news_sheet.append_rows(rows_keep)
                
    except Exception as e:
        log_to_sheet("ERROR", f"封存失敗: {e}")

# --- 爬蟲核心 (V25.0 邏輯回歸) ---

def init_driver():
    chrome_options = Options()
    
    # 這裡恢復了 V23 的開關功能
    if HEADLESS_MODE:
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

        # 1. 填入搜尋框 (互斥邏輯)
        if search_type == "name":
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "tenderName")))
            driver.find_element(By.NAME, "orgName").clear()
        else:
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "orgName")))
            driver.find_element(By.NAME, "tenderName").clear()
            
        input_box.clear()
        input_box.send_keys(keyword)
        time.sleep(0.5) 
        
        # --- ★ V25.0 核心回歸：強制執行 JS ---
        try:
            driver.execute_script("basicTenderSearch();")
        except:
            # 備案：Enter
            input_box.send_keys(Keys.ENTER)
        
        # 3. 等待結果 & 嚴格過濾
        try:
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tb_01")))
            page_text = driver.find_element(By.TAG_NAME, "body").text
            if "無符合條件資料" in page_text or "無資料" in page_text:
                return []
        except:
            return []
        
        # 4. 抓取資料
        results = []
        rows = driver.find_elements(By.CSS_SELECTOR, ".tb_01 tbody tr")
        JUNK_TITLES = ["標案查詢", "決標查詢", "全文檢索", "公告日期查詢", "機關名稱查詢", "功能選項", "更正公告"]

        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 7: continue
            try:
                org_name = cols[1].text.strip()
                date_str = cols[6].text.strip()
                deadline = cols[7].text.strip() if len(cols) > 7 else ""
                budget = cols[8].text.strip() if len(cols) > 8 else ""
                
                links_in_cell = cols[2].find_elements(By.TAG_NAME, "a")
                
                tender_name = ""
                tender_link = ""
                
                # V25 智慧標題抓取 (找最長字串)
                if links_in_cell:
                    longest_link = max(links_in_cell, key=lambda x: len(x.text.strip()))
                    tender_name = longest_link.text.strip()
                    tender_link = longest_link.get_attribute("href")
                else:
                    tender_name = cols[2].text.strip()

                if not tender_name or len(tender_name) < 2: continue
                if any(junk in tender_name for junk in JUNK_TITLES): continue

                results.append({
                    "Date": date_str,
                    "Org": org_name,
                    "Title": tender_name,
                    "Link": tender_link,
                    "Deadline": deadline,
                    "Budget": budget,
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
    new_data_for_notify = [] # ★ 新增這個變數用來存給機器人的資料

    for index, row in df.iterrows():
        if str(row['Link']) not in existing_links:
            row_data = [
                row['Date'], row['Org'], row['Title'], row['Link'],
                row['Deadline'], row['Budget'], row['Tags'], row['Source']
            ]
            new_rows.append(row_data)
            new_data_for_notify.append(row) # ★ 收集新資料
            existing_links.add(str(row['Link']))
    
    if new_rows:
        sheet.append_rows(new_rows)
        # ★ 修改這裡：多回傳一個 pd.DataFrame(new_data_for_notify)
        return len(new_rows), pd.DataFrame(new_data_for_notify)
    
    # ★ 修改這裡：沒資料時回傳 0 和 空DataFrame
    return 0, pd.DataFrame()
    
def load_keywords_from_sheet():
    """讀取雲端關鍵字 (補回遺失的函式)"""
    try:
        client = get_google_client()
        sheet = client.open_by_url(SHEET_URL).worksheet(CONFIG_SHEET_NAME)
        records = sheet.get_all_records()
        keywords = [r['Keyword'] for r in records if r['Type'] == '標案' and r['Keyword']]
        orgs = [r['Keyword'] for r in records if r['Type'] == '機關' and r['Keyword']]
        
        if not keywords: keywords = KEYWORDS
        if not orgs: orgs = ORG_KEYWORDS
        return keywords, orgs
    except:
        return KEYWORDS, ORG_KEYWORDS    

def main():
    print("🚀 啟動爬蟲 (V31.0 Google Chat + 預告戰情版)...")
    
    try:
        keywords, org_keywords = load_keywords_from_sheet()
        driver = init_driver()
        all_data = []
        
        # 1. 爬取「正式公告」
        print("\n--- 1. 搜尋正式公告 ---")
        for org in org_keywords:
            all_data.extend(search_tender(driver, org, "org"))
            time.sleep(1)
        for kw in keywords:
            all_data.extend(search_tender(driver, kw, "name"))
            time.sleep(1)

        # 2. 爬取「採購預告」
        print("\n--- 2. 搜尋採購預告 (Market Intelligence) ---")
        for org in org_keywords:
            all_data.extend(search_forecast(driver, org, "org"))
            time.sleep(1)
            
        driver.quit()
        
        msg = "今日無新情報"
        if all_data:
            df = pd.DataFrame(all_data)
            df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
            
            # 接收兩個回傳值 (數量, 新資料表)
            count, new_df = upload_to_gsheet(df)
            
            if count > 0:
                msg = f"成功執行，發現 {count} 筆新情報 (含預告)"
                # Google Chat 推播
                send_google_chat(count, new_df)
            else:
                msg = "資料已存在 (無新增)"
            
            print(msg)
        
        log_to_sheet("SUCCESS", msg)

    except Exception as e:
        error_msg = f"程式崩潰: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        log_to_sheet("ERROR", error_msg)
        
        # 錯誤通知
        if GOOGLE_CHAT_WEBHOOK:
            try:
                requests.post(GOOGLE_CHAT_WEBHOOK, json={"text": f"🚨 **爬蟲發生錯誤** 🚨\n{str(e)}"})
            except:
                pass
            
        sys.exit(1)

if __name__ == "__main__":
    main()








