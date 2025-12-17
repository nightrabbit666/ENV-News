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

# ★ 設定預算門檻 (單位：元)，低於此金額不存入資料庫
# 設為 0 代表全部都抓；設為 1000000 代表只抓一百萬以上的案子
MIN_BUDGET = 1000000 

KEYWORDS = ["資源回收", "分選", "細分選場", "細分選廠", "細分類", "廢棄物"]
ORG_KEYWORDS = ["資源循環署", "環境管理署"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_KEY_FILE = os.path.join(BASE_DIR, 'key.json')
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit' 
WORKSHEET_NAME = 'news'
LOG_SHEET_NAME = 'logs'
CONFIG_SHEET_NAME = 'Config'
HISTORY_SHEET_NAME = 'history' # 歷史資料分頁
# --- 插入這段 TASKS 設定 ---
# ★ 任務設定：定義雙軌邏輯
TASKS = {
    "General": {
        "config_sheet": "Config",
        "target_sheet": "news",
        "title": "標案戰情快訊",
        "mode": "general" # 一般模式：全部混搜
    },
    "Enterprise": {
        "config_sheet": "Enterprise_Config",
        "target_sheet": "enterprise_news",
        "title": "【企專】標案快訊",
        "mode": "enterprise" # 企專模式：關鍵字優先 + 機關自動分類
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

def load_keywords_from_sheet(sheet_name): # <--- 這裡加了參數
    try:
        client = get_google_client()
        sheet = client.open_by_url(SHEET_URL).worksheet(sheet_name) # <---這裡改成用參數
        records = sheet.get_all_records()
        kws = [r['Keyword'] for r in records if r['Type'] == '標案' and r['Keyword']]
        orgs = [r['Keyword'] for r in records if r['Type'] == '機關' and r['Keyword']]
        return (kws if kws else KEYWORDS), (orgs if orgs else ORG_KEYWORDS)
    except:
        print(f"⚠️ 無法讀取設定檔: {sheet_name}")
        return [], []

# 輔助：解析預算金額
def parse_budget(budget_str):
    try:
        return int(re.sub(r'[^\d]', '', budget_str))
    except:
        return 0

# --- 自動封存舊資料 ---
def archive_old_records():
    print("\n📦 檢查資料封存 (Archive)...")
    try:
        client = get_google_client()
        doc = client.open_by_url(SHEET_URL)
        news_sheet = doc.worksheet(WORKSHEET_NAME)
        
        try:
            history_sheet = doc.worksheet(HISTORY_SHEET_NAME)
        except:
            print("❌ 找不到 history 分頁，跳過封存")
            return

        all_records = news_sheet.get_all_records()
        if not all_records: return

        # 設定保留天數 (超過 90 天就移入歷史區)
        deadline = datetime.now() - timedelta(days=90)
        rows_keep = []
        rows_archive = []
        header = news_sheet.row_values(1) # 保留標題列

        for row in all_records:
            try:
                # 處理日期格式 (民國年轉西元)
                d_str = str(row['Date'])
                parts = d_str.split('/')
                row_date = datetime(int(parts[0]) + 1911, int(parts[1]), int(parts[2]))
                
                if row_date < deadline:
                    rows_archive.append(list(row.values()))
                else:
                    rows_keep.append(list(row.values()))
            except:
                rows_keep.append(list(row.values()))

        if rows_archive:
            print(f"   -> 移動 {len(rows_archive)} 筆舊資料至 history...")
            history_sheet.append_rows(rows_archive)
            news_sheet.clear()
            news_sheet.append_row(header)
            if rows_keep:
                news_sheet.append_rows(rows_keep)
        else:
            print("   -> 無需封存")

    except Exception as e:
        print(f"❌ 封存失敗: {e}")

# --- Google Chat 推播 ---
def send_google_chat(new_data_count, df_new):
    if not GOOGLE_CHAT_WEBHOOK: return
    print("📲 發送 Google Chat 通知...")
    today = datetime.now().strftime("%Y/%m/%d")
    
    text = f"🔔 *【標案戰情快訊】 {today}*\n"
    if new_data_count == 0:
        text += "☕ 今日無新資料 (或未達金額門檻)\n━━━━━━━━━━━━━━\n"
    else:
        text += f"發現 {new_data_count} 筆新商機：\n━━━━━━━━━━━━━━\n"
        count = 0
        for index, row in df_new.iterrows():
            count += 1
            if count > 15:
                text += f"\n...(略 {new_data_count - 15} 筆)"
                break
            
            title = str(row['Title'])
            display_title = title[:30] + "..." if len(title) > 30 else title
            
            text += f"{count}. [{row['Org']}] {row['Org']}\n"
            text += f"   📝 {display_title}\n"
            if row['Budget']: text += f"   💰 {row['Budget']}\n"
            text += f"   ⏳ 截止: {row['Deadline']}\n"
            text += f"   🔗 <{row['Link']}|查看公告> | 📊 <{DASHBOARD_URL}|戰情儀表板>\n\n"

    try:
        requests.post(GOOGLE_CHAT_WEBHOOK, json={"text": text})
    except: pass

# --- 爬蟲核心 ---
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
                    "Tags": f"公告-{keyword}",
                    "Source": "政府採購網"
                })
            except: continue
        return results
    except: return []

def upload_to_gsheet(df, sheet_name): # <--- 這裡加了參數
    print(f"\n☁️ 上傳至 {sheet_name}...")
    client = get_google_client()
    sheet = client.open_by_url(SHEET_URL).worksheet(sheet_name) # <---這裡改成用參數
    existing_data = sheet.get_all_records()
    existing_links = set(str(row['Link']) for row in existing_data if 'Link' in row)
    
    new_rows = []
    new_data_for_notify = []

    for index, row in df.iterrows():
        # ★ 預算過濾器
        budget_val = parse_budget(row['Budget'])
        if MIN_BUDGET > 0 and budget_val < MIN_BUDGET:
            continue

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

# --- 替換整個 main 函式 ---
def main():
    print("🚀 啟動爬蟲 (雙軌分類版)...")
    try:
        driver = init_driver()
        
        # 迴圈執行 TASKS 中的每一個任務
        for task_name, config in TASKS.items():
            print(f"\n======== 執行任務：{task_name} ========")
            
            # 1. 傳入分頁名稱讀取設定
            keywords, org_keywords = load_keywords_from_sheet(config['config_sheet'])
            
            if not keywords: 
                print("   ⚠️ 無關鍵字，跳過")
                continue

            all_data = []

            # 2. 搜尋邏輯
            if config['mode'] == "general":
                # 一般模式：全部混搜 (原本的邏輯)
                print("   [一般模式] 搜尋機關與關鍵字...")
                for org in org_keywords:
                    res = search_tender(driver, org, "org")
                    for r in res: r['Tags'] = f"機關-{org}" # 簡單標記
                    all_data.extend(res)
                    time.sleep(1)
                for kw in keywords:
                    res = search_tender(driver, kw, "name")
                    for r in res: r['Tags'] = f"標案-{kw}"
                    all_data.extend(res)
                    time.sleep(1)

            elif config['mode'] == "enterprise":
                # 企專模式：只搜「標案名稱」，抓回來後再用程式分類
                print("   [企專模式] 搜尋關鍵字並進行機關分類...")
                for kw in keywords:
                    res = search_tender(driver, kw, "name")
                    for r in res:
                        # ★ 自動分類：檢查機關是否在「重點清單」中
                        # 使用 any() 檢查此標案的 Org 是否包含 org_keywords 裡的任一字串
                        is_target = any(target in r['Org'] for target in org_keywords)
                        
                        if is_target:
                            r['Tags'] = "★重點" # 前端網頁會抓這個標記來分組
                        else:
                            r['Tags'] = "其他"
                    
                    all_data.extend(res)
                    time.sleep(1)

            # 3. 存檔與通知
            if all_data:
                df = pd.DataFrame(all_data)
                df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
                
                # 傳入目標分頁名稱
                count, new_df = upload_to_gsheet(df, config['target_sheet'])
                
                # 發送通知 (標題帶入任務名稱)
                if count > 0:
                    title_text = f"{config['title']} (新增 {count} 筆)"
                    send_google_chat(count, new_df) # 這裡會共用同一個 Webhook
                else:
                    print(f"   -> {task_name} 無新資料")
            else:
                print(f"   -> {task_name} 搜尋無結果")

        # 任務結束，記錄日誌
        log_to_sheet("SUCCESS", "雙軌任務執行完畢")

    except Exception as e:
        error_msg = f"程式崩潰: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        log_to_sheet("ERROR", error_msg)
    finally:
        try: driver.quit()
        except: pass

if __name__ == "__main__":
    main()



