import time
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

# --- 設定 ---
# 1. 關鍵字設定
KEYWORDS = ["資源回收", "分選", "細分選場", "細分選廠", "細分類", "廢棄物"]
ORG_KEYWORDS = ["資源循環署", "環境管理署"]

# 2. Google Sheets 設定
JSON_KEY_FILE = 'key.json'  # 請確認這個檔案在同資料夾
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1oJlYFwsipBg1hGMuUYuOWen2jlX19MDJomukvEoahUE/edit?gid=0#gid=0' # ⚠️ 請換成您的試算表網址
WORKSHEET_NAME = 'news'     # 請確認工作表名稱正確

# 目標網址 (PIS)
TARGET_URL = "https://web.pcc.gov.tw/pis/"

def init_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless") # <--- 請把前面的 # 拿掉，一定要開啟！
    chrome_options.add_argument("--window-size=1280,800")
    # ...其他不變
def search_pis(driver, keyword):
    print(f"\n🔍 [PIS] 正在搜尋：{keyword} ...")
    results = []
    try:
        driver.get(TARGET_URL)
        wait = WebDriverWait(driver, 15)
        
        # 搜尋
        try:
            input_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='text']")))
            input_box.clear()
            input_box.send_keys(keyword)
            time.sleep(0.5)
            input_box.send_keys(Keys.ENTER)
        except:
            return []

        # 等待結果
        try:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "a")))
            time.sleep(2)
        except:
            return []
            
        # 抓取資料
        links_elements = driver.find_elements(By.CSS_SELECTOR, "a[href*='tender']")
        date_str = datetime.now().strftime("%Y-%m-%d")

        for elem in links_elements:
            try:
                title = elem.text.strip()
                link = elem.get_attribute("href")
                if len(title) < 4: continue
                
                if not any(d['Link'] == link for d in results):
                    # 這裡整理成 Dictionary 方便 DataFrame 處理
                    results.append({
                        "Date": date_str,
                        "Tags": f"PIS搜尋-{keyword}",
                        "Title": title,
                        "Link": link,
                        "Source": "政府採購網PIS"
                    })
                if len(results) >= 10: break
            except: continue
        return results
    except Exception as e:
        print(f"   ❌ 錯誤: {e}")
        return []

def upload_to_gsheet(df):
    print("\n☁️ 正在連線 Google Sheets...")
    try:
        # 連線設定
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
        client = gspread.authorize(creds)
        
        # 開啟試算表
        sheet = client.open_by_url(SHEET_URL).worksheet(WORKSHEET_NAME)
        
        # 讀取現有資料 (為了防重複)
        # 假設 Link 在第 4 欄 (D欄)，Python index 是 3
        # 我們直接讀取整張表
        existing_data = sheet.get_all_records()
        existing_links = set(row['Link'] for row in existing_data if 'Link' in row)
        
        # 過濾新資料
        new_rows = []
        for index, row in df.iterrows():
            if row['Link'] not in existing_links:
                # 轉成 List 格式準備寫入: [Date, Tags, Title, Link, Source]
                # 順序必須跟 Google Sheet 欄位順序一樣！
                row_data = [
                    row['Date'],
                    row['Tags'],
                    row['Title'],
                    row['Link'],
                    row['Source']
                ]
                new_rows.append(row_data)
                existing_links.add(row['Link']) # 避免本次批次內重複
        
        # 寫入
        if new_rows:
            sheet.append_rows(new_rows)
            print(f"✅ 成功上傳 {len(new_rows)} 筆新資料到雲端！")
        else:
            print("⚠️ 沒有新的不重複資料需上傳。")
            
    except Exception as e:
        print(f"❌ 上傳 Google Sheets 失敗: {e}")
        print("   (請檢查 key.json 是否存在、Email 是否已加入共用)")

def main():
    driver = init_driver()
    all_data = []
    try:
        search_list = KEYWORDS + ORG_KEYWORDS
        for kw in search_list:
            data = search_pis(driver, kw)
            all_data.extend(data)
            time.sleep(1)
    finally:
        driver.quit()
        
    if all_data:
        # 1. 轉成 DataFrame
        df = pd.DataFrame(all_data)
        df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
        
        # 2. 存本機 Excel (備份)
        filename = f"pis_tenders_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        df.to_excel(filename, index=False)
        print(f"\n✅ Excel 已儲存：{filename}")
        
        # 3. 上傳 Google Sheets
        upload_to_gsheet(df)
        
    else:
        print("❌ 沒抓到資料")

if __name__ == "__main__":

    main()
