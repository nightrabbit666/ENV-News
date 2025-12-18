import time
import os
import sys
import pandas as pd
import gspread
import traceback
import requests
import re
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

# --- 設定區 ---
HEADLESS_MODE = True
GOOGLE_CHAT_WEBHOOK = "https://chat.googleapis.com/v1/spaces/AAQADqt_uZc/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=MOuLzkHQFpQP_RDAlmdhWIzw3eWcl6xkUX5_WU09kzw"

# 預算門檻
MIN_BUDGET = 1000000 

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 

# 分頁名稱
LOG_SHEET_NAME = 'logs'

# ★ 任務設定：定義雙軌邏輯 (請確認 Sheet 名稱是否正確)
TASKS = {
    "General": {
        "config_sheet": "Config",
        "target_sheet": "news",
        "title": "標案戰情快訊",
        "mode": "general" 
    },
    "Enterprise": {
        "config_sheet": "Enterprise_Config",
        "target_sheet": "enterprise_news",
        "title": "【企專】標案快訊",
        "mode": "enterprise"
    }
}

URL_BASIC = "https://web.pcc.gov.tw/prkms/tender/common/basic/indexTenderBasic"
DASHBOARD_URL = "https://nightrabbit666.github.io/ENV-News/index.html"

# --- 基礎建設 ---

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
    except: pass

def load_keywords_from_sheet(sheet_name):
    print(f"📖 讀取設定檔: {sheet_name}...")
    try:
        client = get_google_client()
        sheet = client.open_by_url(SHEET_URL).worksheet(sheet_name)
        records = sheet.get_all_records()
        kws = [r['Keyword'] for r in records if r['Type'] == '標案' and r['Keyword']]
        orgs = [r['Keyword'] for r in records if r['Type'] == '機關' and r['Keyword']]
        return kws, orgs
    except Exception as e:
        print(f"⚠️ 讀取失敗 ({e})，使用空列表")
        return [], []

def parse_budget(budget_str):
    try:
        return int(re.sub(r'[^\d]', '', budget_str))
    except:
        return 0

# --- 核心邏輯 ---

def init_driver():
    chrome_options = Options()
    if HEADLESS_MODE: chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
    try:
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e: raise Exception(f"瀏覽器啟動失敗: {e}")

def search_tender(driver, keyword, search_type):
    print(f"   🔍 搜尋 {search_type}：{keyword}")
    try:
        driver.get(URL_BASIC)
        wait = WebDriverWait(driver, 10)
        if search_type == "name":
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "tenderName")))
            driver.find_element(By.NAME, "orgName").clear()
        else:
            input_box = wait.until(EC.visibility_of_element_located((By.NAME, "orgName")))
            driver.find_element(By.NAME, "tenderName").clear()
        input_box.clear()
        input_box.send_keys(keyword)
        try: driver.execute_script("basicTenderSearch();")
        except: input_box.send_keys(Keys.ENTER)
        
        try:
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "tb_01")))
            if "無符合條件資料" in driver.page_source: return []
        except: return []
        
        results = []
        rows = driver.find_elements(By.CSS_SELECTOR, ".tb_01 tbody tr")
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 7: continue
            try:
                tender_name = cols[2].text.strip()
                link = cols[2].find_elements(By.TAG_NAME, "a")[0].get_attribute("href") if cols[2].find_elements(By.TAG_NAME, "a") else ""
                if not tender_name: continue
                results.append({
                    "Date": cols[6].text.strip(),
                    "Org": cols[1].text.strip(),
                    "Title": tender_name,
                    "Link": link,
                    "Deadline": cols[7].text.strip(),
                    "Budget": cols[8].text.strip(),
                    "Tags": "", # 稍後填入
                    "Source": "政府採購網"
                })
            except: continue
        return results
    except: return []

def upload_to_gsheet(df, sheet_name):
    print(f"☁️ 上傳至 {sheet_name}...")
    client = get_google_client()
    sheet = client.open_by_url(SHEET_URL).worksheet(sheet_name)
    existing_data = sheet.get_all_records()
    existing_links = set(str(row['Link']) for row in existing_data if 'Link' in row)
    
    new_rows = []
    new_data_for_notify = []

    for index, row in df.iterrows():
        budget_val = parse_budget(row['Budget'])
        if MIN_BUDGET > 0 and budget_val < MIN_BUDGET: continue

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

# --- Google Chat 推播 (分頁版：突破字數限制) ---
def send_google_chat(new_data_count, df_new):
    if not GOOGLE_CHAT_WEBHOOK: return
    print("📲 發送 Google Chat 通知...")
    today = datetime.now().strftime("%Y/%m/%d")
    
    # 如果沒資料，發送一則簡單通知就好
    if new_data_count == 0:
        text = f"🔔 *【標案戰情快訊】 {today}*\n"
        text += "☕ 今日無新資料 (或未達金額門檻)\n━━━━━━━━━━━━━━\n"
        try: requests.post(GOOGLE_CHAT_WEBHOOK, json={"text": text})
        except: pass
        return

    # ★ 設定每則訊息最多顯示幾筆 (建議 20 筆，避免超過 Google 4096 字元限制)
    BATCH_SIZE = 20
    
    # 將資料轉為列表方便切分
    records = df_new.to_dict('records')
    total_batches = (len(records) + BATCH_SIZE - 1) // BATCH_SIZE  # 計算總共要發幾則

    for i in range(0, len(records), BATCH_SIZE):
        batch_data = records[i : i + BATCH_SIZE]
        current_batch_num = (i // BATCH_SIZE) + 1
        
        # 標題加上 (1/3) 這種頁碼，讓你知道還有下一則
        header = f"🔔 *【標案戰情快訊】 {today}* ({current_batch_num}/{total_batches})\n"
        header += f"發現 {new_data_count} 筆新商機：\n━━━━━━━━━━━━━━\n"
        
        text = header
        for idx, row in enumerate(batch_data):
            # 全局序號 (例如第 21 筆)
            global_idx = i + idx + 1
            
            title = str(row['Title'])
            # 標題過長稍微截斷，避免佔用太多字數
            display_title = title[:35] + "..." if len(title) > 35 else title
            
            text += f"{global_idx}. [{row['Org']}] {row['Org']}\n"
            text += f"   📝 {display_title}\n"
            if row['Budget']: text += f"   💰 {row['Budget']}\n"
            text += f"   ⏳ 截止: {row['Deadline']}\n"
            text += f"   🔗 <{row['Link']}|查看公告> | 📊 <{DASHBOARD_URL}|戰情儀表板>\n\n"

        # 發送這一批
        try:
            requests.post(GOOGLE_CHAT_WEBHOOK, json={"text": text})
            time.sleep(0.5) # 稍微休息一下，避免發送太快順序錯亂
        except Exception as e:
            print(f"❌ 發送失敗: {e}")

# --- 主程式 ---
def main():
    print("🚀 啟動爬蟲 (雙軌分類版)...")
    driver = init_driver()
    
    try:
        for task_name, config in TASKS.items():
            print(f"\n======== 執行任務：{task_name} ========")
            
            # 1. 讀取設定
            keywords, org_keywords = load_keywords_from_sheet(config['config_sheet'])
            
            # 2. 執行搜尋邏輯 (依模式區分)
            all_data = []
            
            if config['mode'] == "general":
                # [一般模式]：標案名 OR 機關名 (全部混搜)
                for org in org_keywords:
                    res = search_tender(driver, org, "org")
                    for r in res: r['Tags'] = f"機關-{org}"
                    all_data.extend(res)
                    time.sleep(1)
                for kw in keywords:
                    res = search_tender(driver, kw, "name")
                    for r in res: r['Tags'] = f"標案-{kw}"
                    all_data.extend(res)
                    time.sleep(1)
            
            elif config['mode'] == "enterprise":
                # [企專模式]：關鍵字優先搜尋，再依據機關分類
                for kw in keywords:
                    res = search_tender(driver, kw, "name")
                    for r in res:
                        # ★ 核心分類邏輯：這是網頁分組的關鍵
                        is_target_org = any(target in r['Org'] for target in org_keywords)
                        if is_target_org:
                            r['Tags'] = "★重點" 
                        else:
                            r['Tags'] = "其他"
                    
                    all_data.extend(res)
                    time.sleep(1)

            # 3. 處理結果
            if all_data:
                df = pd.DataFrame(all_data)
                df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
                count, new_df = upload_to_gsheet(df, config['target_sheet'])
                send_google_chat(count, new_df, config['title'])
                print(f"   ✅ {task_name} 完成：新增 {count} 筆")
            else:
                print(f"   ✅ {task_name} 完成：無資料")

        log_to_sheet("SUCCESS", "雙軌任務執行完畢")

    except Exception as e:
        error_msg = f"程式崩潰: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        log_to_sheet("ERROR", error_msg)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()

