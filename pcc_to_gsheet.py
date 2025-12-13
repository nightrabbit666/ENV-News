import time
import os
import sys
import pandas as pd
import gspread
import traceback
import requests
import json
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

# --- 【切換開關】 ---
HEADLESS_MODE = True

# --- 設定區 ---
# Google Chat Webhook
GOOGLE_CHAT_WEBHOOK = "https://chat.googleapis.com/v1/spaces/AAQAbfa7gJQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=N4OegGZLJ2y1ANxt41jIFf57RaGV4TI3Vw_GyHzdzeU"

# 預設關鍵字
KEYWORDS = ["資源回收", "分選", "細分選場", "細分選廠", "細分類", "廢棄物"]
ORG_KEYWORDS = ["資源循環署", "環境管理署"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 
WORKSHEET_NAME = 'news'
LOG_SHEET_NAME = 'logs'
CONFIG_SHEET_NAME = 'Config'

# 網址定義
URL_BASIC = "https://web.pcc.gov.tw/prkms/tender/common/basic/indexTenderBasic"   # 正式公告
DASHBOARD_URL = "https://nightrabbit666.github.io/ENV-News/index.html"

# --- 基礎建設函式 ---

def get_google_client():
    if not os.path.exists(JSON_KEY_FILE):
        raise FileNotFoundError(f"找不到 key.json")
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    return gspread.authorize(creds)

def log_to_sheet(status, message):
    print(f"[{status}] {message}")
    try:
        client = get_google_client()
        sheet = client.open_by_url(SHEET_URL).worksheet(LOG_SHEET_NAME)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sheet.append_row([timestamp, status, message])
    except:
        pass

def load_keywords_from_sheet():
    try:
        client = get_google_client()
        sheet = client.open_by_url(SHEET_URL).worksheet(CONFIG_SHEET_NAME)
        records = sheet.get_all_records()
        kws = [r['Keyword'] for r in records if r['Type'] == '標案' and r['Keyword']]
        orgs = [r['Keyword'] for r in records if r['Type'] == '機關' and r['Keyword']]
        
        if not kws: kws = KEYWORDS
        if not orgs: orgs = ORG_KEYWORDS
        return kws, orgs
    except:
        return KEYWORDS, ORG_KEYWORDS

# --- Google Chat 推播 (V31.2 修改版) ---
def send_google_chat(new_data_count, df_new):
    """發送 Google Chat 通知 (必定發送)"""
    if not GOOGLE_CHAT_WEBHOOK: return

    print("📲 準備發送 Google Chat 通知...")
    today = datetime.now().strftime("%Y/%m/%d")
    
    text = f"🔔 *【標案戰情快訊】 {today}*\n"
    
    if new_data_count == 0:
        # 若無資料，顯示這段
        text += "☕ 今日無新資料 (或皆已建檔)\n"
        text += "━━━━━━━━━━━━━━\n"
    else:
        # 若有資料，顯示列表
        text += f"發現 {new_data_count} 筆新商機：\n"
        text += "━━━━━━━━━━━━━━\n"

        count = 0
        for index, row in df_new.iterrows():
            count += 1
            if count > 15:
                text += f"\n...(還有 {new_data_count - 15} 筆，請至儀表板查看)"
                break
            
            title = str(row['Title'])
            if len(title) > 30: title = title[:30] + "..."
            
            text += f"{count}. [{row['Org']}] {row['Org']}\n"
            text += f"   📝 {title}\n"
            if row['Budget']:
                text += f"   💰 {row['Budget']}\n"
            text += f"   ⏳ 截止: {row['Deadline']}\n"
            text += f"   🔗 <{row['Link']}|點擊查看>\n\n"

    text += "━━━━━━━━━━━━━━\n"
    text += f"📊 <{DASHBOARD_URL}|查看完整戰情儀表板>"

    try:
        response = requests.post(GOOGLE_CHAT_WEBHOOK, json={"text": text})
        if response.status_code == 200:
            print("✅ Google Chat 發送成功！")
        else:
            print(f"❌ Google Chat 發送失敗: {response.text}")
    except Exception as e:
        print(f"❌ Google Chat 連線錯誤: {e}")

# --- 爬蟲核心 ---

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
        return webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        raise Exception(f"瀏覽器啟動失敗: {e}")

# 正式公告爬蟲
def search_tender(driver, keyword, search_type):
    print(f"\n🔍 [公告] 搜尋 {search_type}：{keyword}")
    try:
        driver.get(URL_BASIC)
        wait = WebDriverWait(driver, 15)

        if search_type == "name":
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "tenderName")))
            driver.find_element(By.NAME, "orgName").clear()
        else:
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "orgName")))
            driver.find_element(By.NAME, "tenderName").clear()
            
        input_box.clear()
        input_box.send_keys(keyword)
        
        try:
            driver.execute_script("basicTenderSearch();")
        except:
            input_box.send_keys(Keys.ENTER)
        
        try:
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tb_01")))
            if "無符合條件資料" in driver.page_source: return []
        except: return []
        
        results = []
        rows = driver.find_elements(By.CSS_SELECTOR, ".tb_01 tbody tr")
        JUNK_TITLES = ["標案查詢", "決標查詢", "全文檢索"]

        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 7: continue
            try:
                tender_name = cols[2].text.strip()
                link = ""
                links = cols[2].find_elements(By.TAG_NAME, "a")
                if links:
                    link = links[0].get_attribute("href")
                    tender_name = links[0].text.strip()

                if not tender_name or any(j in tender_name for j in JUNK_TITLES): continue

                results.append({
                    "Date": cols[6].text.strip(),
                    "Org": cols[1].text.strip(),
                    "Title": tender_name,
                    "Link": link,
                    "Deadline": cols[7].text.strip(),
                    "Budget": cols[8].text.strip(),
                    "Tags": f"公告-{keyword}",
                    "Source": "政府採購網"
                })
            except: continue
        return results
    except Exception as e:
        print(f"   ❌ 公告搜尋錯誤: {e}")
        return []

def upload_to_gsheet(df):
    print("\n☁️ 上傳 Google Sheets...")
    client = get_google_client()
    sheet = client.open_by_url(SHEET_URL).worksheet(WORKSHEET_NAME)
    existing_data = sheet.get_all_records()
    existing_links = set(str(row['Link']) for row in existing_data if 'Link' in row)
    
    new_rows = []
    new_data_for_notify = []

    for index, row in df.iterrows():
        if str(row['Link']) not in existing_links:
            row_data = [
                row['Date'], row['Org'], row['Title'], row['Link'],
                row['Deadline'], row['Budget'], row['Tags'], row['Source']
            ]
            new_rows.append(row_data)
            new_data_for_notify.append(row)
            existing_links.add(str(row['Link']))
    
    if new_rows:
        sheet.append_rows(new_rows)
        return len(new_rows), pd.DataFrame(new_data_for_notify)
    
    return 0, pd.DataFrame()

def main():
    print("🚀 啟動爬蟲 (V31.2 Google Chat 必定通知版)...")
    
    try:
        keywords, org_keywords = load_keywords_from_sheet()
        driver = init_driver()
        all_data = []
        
        print("\n--- 搜尋正式公告 ---")
        for org in org_keywords:
            all_data.extend(search_tender(driver, org, "org"))
            time.sleep(1)
        for kw in keywords:
            all_data.extend(search_tender(driver, kw, "name"))
            time.sleep(1)

        driver.quit()
        
        msg = "今日無新情報"
        count = 0
        new_df = pd.DataFrame()

        if all_data:
            df = pd.DataFrame(all_data)
            df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
            
            # 上傳並取得新資料
            count, new_df = upload_to_gsheet(df)
            
            if count > 0:
                msg = f"成功執行，發現 {count} 筆新情報"
            else:
                msg = "資料已存在 (無新增)"
        else:
            msg = "搜尋無結果"
        
        # ★ 移到 if 外面：無論 count 是多少，都發送通知
        send_google_chat(count, new_df)
        
        print(msg)
        log_to_sheet("SUCCESS", msg)

    except Exception as e:
        error_msg = f"程式崩潰: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        log_to_sheet("ERROR", error_msg)
        
        if GOOGLE_CHAT_WEBHOOK:
            try:
                requests.post(GOOGLE_CHAT_WEBHOOK, json={"text": f"🚨 **爬蟲發生錯誤** 🚨\n{str(e)}"})
            except: pass
            
        sys.exit(1)

if __name__ == "__main__":
    main()
