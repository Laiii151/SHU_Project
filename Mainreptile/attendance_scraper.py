# -*- coding: utf-8 -*-
"""
世新大學 缺勤記錄爬蟲
從世新校網進入學生教務系統，爬取個人缺勤記錄
"""

import time
import os
from typing import List, Tuple, Optional, Dict, Any
import re
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager

# 載入環境變數
try:
    from dotenv import load_dotenv
    load_dotenv()
    print("✅ 已載入 .env 檔案")
except ImportError:
    print("⚠️ 未安裝 python-dotenv，請執行: pip install python-dotenv")
    print("⚠️ 或手動設定環境變數")

# ========= 設定區 =========
# 從環境變數讀取帳號密碼
USERNAME = os.getenv('SHU_USERNAME')
PASSWORD = os.getenv('SHU_PASSWORD')

# 檢查是否有設定帳號密碼
if not USERNAME or PASSWORD is None:
    print("❌ 錯誤：請在 .env 檔案中設定 SHU_USERNAME 和 SHU_PASSWORD")
    print("📝 .env 檔案格式範例：")
    print("SHU_USERNAME=你的學號")
    print("SHU_PASSWORD=你的密碼")
    exit(1)

HOME_URL = "https://www.shu.edu.tw/"
HEADLESS = False
MAX_WAIT = 25

print(f"🔐 使用帳號：{USERNAME[:3]}***{USERNAME[-3:] if len(USERNAME) > 6 else '***'}")

# ---------------- 基礎工具函數 ----------------
def build_driver():
    """建立 Chrome WebDriver"""
    opt = webdriver.ChromeOptions()
    if HEADLESS:
        opt.add_argument("--headless=new")
    opt.add_argument("--no-sandbox")
    opt.add_argument("--disable-gpu")
    opt.add_argument("--disable-blink-features=AutomationControlled")
    opt.add_argument("--window-size=1440,900")
    opt.add_experimental_option("excludeSwitches", ["enable-automation"])
    opt.add_experimental_option('useAutomationExtension', False)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opt)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

def js_click(driver, el):
    """JavaScript 點擊元素"""
    driver.execute_script("""
        const el = arguments[0];
        el.scrollIntoView({block:'center'});
        try{ el.dispatchEvent(new MouseEvent('mouseover',{bubbles:true})); }catch(e){}
        try{ el.dispatchEvent(new MouseEvent('mousedown',{bubbles:true})); }catch(e){}
        try{ el.dispatchEvent(new MouseEvent('mouseup',{bubbles:true})); }catch(e){}
        try{ el.click(); }catch(e){}
    """, el)

def find_and_js_click(driver, selector: str, by="css") -> bool:
    """尋找元素並點擊"""
    try:
        if by == "css":
            el = driver.find_element(By.CSS_SELECTOR, selector)
        else:
            el = driver.find_element(By.XPATH, selector)
        js_click(driver, el)
        return True
    except Exception as e:
        print(f"點擊失敗 ({by}: {selector}): {e}")
        return False

def click_first_working(driver, selectors: List[Tuple[str, str]]) -> bool:
    """嘗試多個選擇器，點擊第一個成功的"""
    for by, sel in selectors:
        if find_and_js_click(driver, sel, by=by):
            print(f"✅ 成功點擊: {by}={sel}")
            return True
    return False

def save_html(driver, path):
    """保存頁面HTML"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)

def _die(driver, msg, png, html):
    """錯誤處理：截圖並保存HTML"""
    driver.save_screenshot(png)
    save_html(driver, html)
    print(f"❌ {msg}")
    print(f"📸 截圖已保存：{png}")
    print(f"📄 HTML已保存：{html}")
    raise RuntimeError(f"{msg}；已存 {png} / {html}")

# ---------------- 導覽函數 ----------------
def goto_student_system_from_home(driver):
    """從首頁進入學生教務系統"""
    print("🌐 正在進入世新大學首頁...")
    driver.get(HOME_URL)
    time.sleep(2)
    
    # 點擊校務系統
    print("🔍 尋找校務系統連結...")
    ok = click_first_working(driver, [
        ("css", "body > div.logosearch-area > div.n2021-area > p > a:nth-child(4)"),
        ("xpath", "//a[contains(@href,'System-info.aspx')]"),
        ("xpath", "//a[contains(text(),'校務系統')]"),
    ])
    if not ok:
        _die(driver, "找不到『校務系統』連結", "fail_sys_link.png", "fail_sys_link.html")
    
    time.sleep(3)
    
    # 點擊學生教務系統
    print("🔍 尋找學生教務系統連結...")
    ok = click_first_working(driver, [
        ("css", "body > div:nth-child(10) > div > div.sm-page-all-area > div:nth-child(2) > div.ct-sub-sbox.ct-sub-nsbox.ct-sub-nsortbox > a:nth-child(9)"),
        ("css", "body > div:nth-child(11) > div > div.sm-page-all-area > div:nth-child(2) > div.ct-sub-sbox.ct-sub-nsbox.ct-sub-nsortbox > a:nth-child(9)"),
        ("xpath", "//a[contains(@href,'stulb.shu.edu.tw')]"),
        ("xpath", "//a[normalize-space()='學生教務系統' or contains(normalize-space(.),'學生教務系統')]"),
    ])
    
    if not ok:
        print("⚠️ 找不到學生教務系統連結，直接開啟網址...")
        driver.execute_script("window.open('https://stulb.shu.edu.tw/','_blank');")
    
    # 切換到新分頁
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(2)

def login_if_needed(driver):
    """如需要則進行登入"""
    print("🔐 檢查是否需要登入...")
    
    try:
        # 等待登入表單出現
        username_field = WebDriverWait(driver, 12).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'],input[autocomplete='username']"))
        )
        password_field = WebDriverWait(driver, 12).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password'],input[autocomplete='current-password']"))
        )
        
        print("📝 輸入帳號密碼...")
        username_field.clear()
        username_field.send_keys(USERNAME)
        
        password_field.clear()
        password_field.send_keys(PASSWORD)
        
        # 提交登入表單
        try:
            submit_btn = driver.find_element(By.CSS_SELECTOR, "input[type='submit'],button[type='submit']")
            js_click(driver, submit_btn)
        except NoSuchElementException:
            password_field.submit()
        
        print("⏳ 等待登入完成...")
        time.sleep(0.8)
        # 提交後短暫輪詢錯誤訊息，若偵測到立即結束
        end = time.time() + 8
        while time.time() < end:
            try:
                # 先看常見的訊息元素
                try:
                    msg = driver.find_element(By.ID, 'lblMessage').text
                except Exception:
                    msg = ''
                low = (msg or '').lower()
                if any(k in low for k in [
                    '登入帳號或密碼錯誤', '輸入帳號或密碼錯誤', '帳號或密碼錯誤',
                    'login failed', 'invalid password', 'authentication failed']):
                    try:
                        driver.save_screenshot('login_error.png')
                        save_html(driver, 'login_error.html')
                    except Exception:
                        pass
                    print('❌ 登入失敗：', msg)
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    import os
                    os._exit(2)

                # 廣泛比對 body 文字
                try:
                    body_text = driver.find_element(By.TAG_NAME, 'body').text
                except Exception:
                    body_text = ''
                lowb = (body_text or '').lower()
                if any(k in lowb for k in [
                    '登入帳號或密碼錯誤', '輸入帳號或密碼錯誤', '帳號或密碼錯誤',
                    'login failed', 'invalid password', 'authentication failed']):
                    try:
                        driver.save_screenshot('login_error.png')
                        save_html(driver, 'login_error.html')
                    except Exception:
                        pass
                    print('❌ 登入失敗（body）：帳號或密碼錯誤')
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    import os
                    os._exit(2)
            except SystemExit:
                raise
            except Exception:
                pass
            time.sleep(0.5)
        
    except TimeoutException:
        print("ℹ️ 沒有找到登入表單，可能已經登入或頁面結構不同")

def navigate_to_attendance(driver):
    """導覽到缺勤記錄頁面"""
    print("🧭 導覽到缺勤記錄頁面...")
    
    # 切換到 main frame（如果存在）
    try:
        driver.switch_to.default_content()
        driver.switch_to.frame("main")
        print("✅ 已切換到 main frame")
    except Exception:
        print("ℹ️ 沒有找到 main frame，繼續使用預設內容")
    
    # 等待頁面載入完成
    WebDriverWait(driver, 20).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    time.sleep(2)
    
    # 保存當前頁面供除錯
    save_html(driver, "navigation_debug.html")
    
    # 嘗試找到所有 label 元素並列出
    try:
        labels = driver.find_elements(By.CSS_SELECTOR, "span.label")
        print(f"🔍 找到 {len(labels)} 個 label 元素：")
        for i, label in enumerate(labels):
            text = label.text.strip()
            print(f"   {i+1}. {text}")
    except Exception as e:
        print(f"列出 label 失敗: {e}")
    
    print("🔍 尋找課務作業選單...")
    
    def click_label(text) -> bool:
        """點擊指定文字的 label"""
        return driver.execute_script("""
            const wanted = arguments[0].trim();
            const els = Array.from(document.querySelectorAll('.label, a, button, span, div'));
            for (const el of els) {
              const t = (el.textContent || '').trim();
              if (t === wanted) {
                el.scrollIntoView({block:'center'});
                try { el.click(); return true; } catch(e){}
                try { el.parentElement?.click(); return true; } catch(e){}
              }
            }
            return false;
        """, text)
    
    # 第一步：點擊課務作業
    success = False
    selectors_step1 = [
        ("xpath", "//span[@class='label' and normalize-space()='課務作業']"),
        ("xpath", "//span[contains(@class,'label') and contains(text(),'課務作業')]"),
        ("css", "span.label"),  # 找到所有 label 後用 JS 點擊
        ("xpath", "//span[contains(text(),'課務作業')]"),
    ]
    
    # 先嘗試直接點擊
    if click_label("課務作業"):
        print("✅ 成功點擊課務作業（JS方式）")
        success = True
    else:
        # 嘗試不同的選擇器
        for by, selector in selectors_step1:
            try:
                if by == "css" and selector == "span.label":
                    # 找到所有 label，逐個檢查
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    for elem in elements:
                        if "課務作業" in elem.text:
                            js_click(driver, elem)
                            print("✅ 成功點擊課務作業")
                            success = True
                            break
                else:
                    element = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH if by == "xpath" else By.CSS_SELECTOR, selector))
                    )
                    js_click(driver, element)
                    print(f"✅ 成功點擊課務作業: {by}={selector}")
                    success = True
                    break
                    
            except Exception as e:
                print(f"嘗試點擊失敗 ({by}: {selector}): {e}")
                continue
            
            if success:
                break
    
    if not success:
        _die(driver, "找不到『課務作業』選單", "fail_menu1.png", "fail_menu1.html")
    
    time.sleep(2)
    
    # 第二步：點擊缺勤記錄子選單
    print("🔍 尋找缺勤記錄子選單...")
    
    # 再次列出可用的選項
    try:
        labels = driver.find_elements(By.CSS_SELECTOR, "span.label, .label, a, button")
        print("🔍 當前可點擊的元素：")
        for i, label in enumerate(labels[:20]):  # 只顯示前20個
            text = label.text.strip()
            if text:
                print(f"   {i+1}. {text}")
    except Exception as e:
        print(f"列出選項失敗: {e}")
    
    # 嘗試點擊缺勤相關選項
    success = False
    attendance_keywords = ["SC0108-出缺勤記錄查詢", "SC0108", "出缺勤記錄查詢", "出缺勤記錄", "缺勤記錄"]
    
    for keyword in attendance_keywords:
        if click_label(keyword):
            print(f"✅ 成功點擊: {keyword}")
            success = True
            break
    
    if not success:
        # 使用你提供的具體選擇器
        specific_selectors = [
            ("css", "#app > div > ul > div > div:nth-child(2) > div > div.bar-menu-items > div:nth-child(1) > span"),
            ("xpath", "//span[@class='label' and contains(text(), 'SC0108')]"),
            ("xpath", "//span[contains(text(), 'SC0108-出缺勤記錄查詢')]"),
        ]
        
        for by, selector in specific_selectors:
            try:
                element = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR if by == "css" else By.XPATH, selector))
                )
                js_click(driver, element)
                print(f"✅ 成功點擊SC0108: {by}={selector}")
                success = True
                break
            except Exception as e:
                print(f"嘗試點擊失敗 ({by}: {selector}): {e}")
                continue
    
    if not success:
        # 手動搜尋所有元素
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, "span, a, button, div")
            for elem in elements:
                text = elem.text.strip()
                if any(keyword in text for keyword in attendance_keywords):
                    print(f"🎯 找到可能的目標: {text}")
                    js_click(driver, elem)
                    success = True
                    break
        except Exception as e:
            print(f"手動搜尋失敗: {e}")
    
    if not success:
        _die(driver, "找不到SC0108-出缺勤記錄查詢選單", "fail_menu2.png", "fail_menu2.html")
    
    time.sleep(3)
    
    time.sleep(3)

def parse_attendance_data(driver):
    """解析缺勤記錄數據（包含滾動載入）"""
    print("📊 開始解析缺勤記錄...")
    
    # 等待頁面載入
    time.sleep(3)
    
    # 保存當前頁面供除錯
    save_html(driver, "attendance_debug.html")
    
    attendance_records = []
    
    try:
        # 先滾動到頂部
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        
        # 滾動載入所有資料
        print("🔄 滾動頁面載入所有資料...")
        last_height = driver.execute_script("return document.body.scrollHeight")
        scroll_attempts = 0
        max_scroll_attempts = 10
        
        while scroll_attempts < max_scroll_attempts:
            # 滾動到頁面底部
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)  # 等待內容載入
            
            # 檢查是否有新內容載入
            new_height = driver.execute_script("return document.body.scrollHeight")
            
            if new_height == last_height:
                # 沒有新內容，可能已經載入完畢
                print(f"   滾動完成，共滾動 {scroll_attempts} 次")
                break
            
            last_height = new_height
            scroll_attempts += 1
            print(f"   第 {scroll_attempts} 次滾動，頁面高度: {new_height}")
        
        # 滾動回頂部開始解析
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        
        # 方法1：嘗試解析表格
        tables = driver.find_elements(By.TAG_NAME, "table")
        
        for i, table in enumerate(tables):
            print(f"🔍 檢查第 {i+1} 個表格...")
            table_text = table.text
            
            # 檢查是否包含缺勤相關內容
            if any(keyword in table_text for keyword in ['學年', '學期', '課程', '缺勤', '出缺席', '曠課', 'SC0108']):
                print(f"✅ 找到缺勤記錄表格 #{i+1}")
                records = parse_attendance_table(table)
                attendance_records.extend(records)
        
        # 方法2：如果沒有找到表格，嘗試解析其他結構
        if not attendance_records:
            print("🔄 嘗試解析非表格結構...")
            records = parse_attendance_text(driver)
            attendance_records.extend(records)
        
        # 方法3：尋找特定的資料容器
        if not attendance_records:
            print("🔄 嘗試尋找特定的資料容器...")
            records = parse_attendance_containers(driver)
            attendance_records.extend(records)
    
    except Exception as e:
        print(f"❌ 解析過程發生錯誤: {e}")
        
        # 保存除錯資訊
        try:
            page_text = driver.find_element(By.TAG_NAME, "body").text
            with open("attendance_page_text.txt", "w", encoding="utf-8") as f:
                f.write(page_text)
            print("📝 已保存頁面文字到 attendance_page_text.txt")
        except:
            pass
    
    return attendance_records

def parse_attendance_containers(driver):
    """解析可能包含缺勤記錄的容器元素"""
    records = []
    
    try:
        # 尋找可能的資料容器
        container_selectors = [
            "div[class*='table']",
            "div[class*='data']", 
            "div[class*='record']",
            "div[class*='content']",
            ".ant-table-tbody tr",  # Ant Design 表格
            "[class*='row']",
            "[data-row]"
        ]
        
        for selector in container_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                if not elements:
                    continue
                
                print(f"🔍 檢查容器: {selector} (找到 {len(elements)} 個)")
                
                for i, elem in enumerate(elements):
                    text = elem.text.strip()
                    
                    # 檢查是否包含課程代碼或相關資訊
                    if any(pattern in text for pattern in ['GENS-', 'INF-', '學年', '第一學期', '第二學期']):
                        record = extract_record_from_text(text)
                        if record:
                            records.append(record)
                            if len(records) % 5 == 0:
                                print(f"   已從容器解析 {len(records)} 筆記錄")
                
                if records:
                    print(f"✅ 從 {selector} 解析到 {len(records)} 筆記錄")
                    break
                    
            except Exception as e:
                print(f"解析容器 {selector} 失敗: {e}")
                continue
                
    except Exception as e:
        print(f"容器解析錯誤: {e}")
    
    return records

def extract_record_from_text(text):
    """從文字中提取單筆記錄"""
    try:
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        record = {}
        
        # 尋找學年學期
        for line in lines:
            if re.search(r'\d{3}.*?第.*?學期', line):
                record['學年學期'] = line
                break
        
        # 尋找課程代碼
        for line in lines:
            if re.search(r'[A-Z]{2,4}-\d{3}-\d{2}-[A-Z]\d', line):
                record['課程代碼'] = line
                break
        
        # 尋找課程名稱
        for line in lines:
            if '通識' in line or any(char in line for char in '：古文英數資管'):
                if '課程代碼' not in record or line != record.get('課程代碼'):
                    record['課程名稱'] = line
                    break
        
        # 尋找教師名稱
        for line in lines:
            if len(line) <= 6 and all(ord(char) > 127 for char in line if char):  # 可能是中文姓名
                if line not in record.values():
                    record['授課教師'] = line
                    break
        
        # 尋找缺勤狀態
        for line in lines:
            if any(keyword in line for keyword in ['不扣考', '扣考', '曠課', '請假', '明細']):
                record['缺勤狀態'] = line
        
        # 如果有有效資料則返回
        if len(record) >= 2:  # 至少要有2個欄位
            return record
            
    except Exception as e:
        print(f"提取記錄失敗: {e}")
    
    return None

def parse_attendance_table(table):
    """解析缺勤記錄表格"""
    records = []
    
    try:
        rows = table.find_elements(By.TAG_NAME, "tr")
        headers = []
        
        for i, row in enumerate(rows):
            cells = row.find_elements(By.TAG_NAME, "td")
            if not cells:
                cells = row.find_elements(By.TAG_NAME, "th")
            
            if not cells:
                continue
            
            cell_texts = [cell.text.strip() for cell in cells]
            
            # 第一行通常是表頭
            if i == 0 or not headers:
                if any('學年' in text or '課程' in text for text in cell_texts):
                    headers = cell_texts
                    print(f"📋 表頭: {headers}")
                    continue
            
            # 解析資料行
            if len(cell_texts) >= len(headers) and headers:
                record = {}
                for j, header in enumerate(headers):
                    if j < len(cell_texts):
                        record[header] = cell_texts[j]
                
                # 檢查是否為有效記錄
                if any(record.values()) and record:
                    records.append(record)
                    if len(records) % 5 == 0:
                        print(f"   已解析 {len(records)} 筆記錄")
    
    except Exception as e:
        print(f"表格解析錯誤: {e}")
    
    return records

def parse_attendance_text(driver):
    """解析頁面文字內容（備用方法）"""
    records = []
    
    try:
        # 取得頁面所有文字
        page_text = driver.find_element(By.TAG_NAME, "body").text
        lines = [line.strip() for line in page_text.split('\n') if line.strip()]
        
        current_record = {}
        
        for line in lines:
            # 根據實際頁面結構調整解析邏輯
            # 這裡需要根據你實際看到的頁面格式來調整
            
            # 學年學期識別
            if re.match(r'\d{3}\s*學年', line):
                if current_record:
                    records.append(current_record)
                current_record = {'學年學期': line}
            
            # 課程資訊識別
            elif '課程' in line or 'GENS-' in line or 'INF-' in line:
                current_record['課程資訊'] = line
            
            # 缺勤狀態識別
            elif any(keyword in line for keyword in ['不扣考', '明細', '曠課', '請假']):
                current_record['缺勤狀態'] = line
        
        # 加入最後一筆記錄
        if current_record:
            records.append(current_record)
    
    except Exception as e:
        print(f"文字解析錯誤: {e}")
    
    return records

def clean_attendance_data(records):
    """清理缺勤記錄資料"""
    if not records:
        return pd.DataFrame()
    
    print(f"🧹 清理前：{len(records)} 筆記錄")
    
    # 轉換為 DataFrame
    df = pd.DataFrame(records)
    
    # 移除空記錄
    df = df.dropna(how='all').reset_index(drop=True)
    
    # 統一欄位名稱
    column_mapping = {
        '學年': '學年',
        '學期': '學期', 
        '課程代碼': '課程代碼',
        '課程名稱': '課程名稱',
        '授課教師': '教師',
        '曠課次數': '曠課次數',
        '扣考時數': '扣考時數',
        '扣考': '扣考狀態',
        '備註': '備註'
    }
    
    # 重新命名欄位
    for old_name, new_name in column_mapping.items():
        if old_name in df.columns:
            df = df.rename(columns={old_name: new_name})
    
    # 欄位順序與對齊（避免顯示時欄位錯位）
    preferred_order = ['學年','學期','課程代碼','課程名稱','教師','缺勤狀態','曠課次數','扣考時數','備註']
    ordered_exist = [c for c in preferred_order if c in df.columns]
    others = [c for c in df.columns if c not in ordered_exist]
    if ordered_exist:
        df = df[ordered_exist + others]

    # 基本修整：去除字串前後空白
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.replace('\u00a0',' ', regex=False).str.strip()

    print(f"🧹 清理後：{len(df)} 筆記錄")
    print(f"📊 欄位: {list(df.columns)}")
    
    return df

# ---------------- 主程式 ----------------
def main():
    """主程式入口"""
    driver = build_driver()
    
    try:
        print("🚀 開始執行缺勤記錄爬蟲...")
        print("=" * 60)
        
        # 步驟1: 進入學生教務系統
        goto_student_system_from_home(driver)
        print("✅ 已進入學生教務系統")
        
        # 步驟2: 登入
        login_if_needed(driver)
        print("✅ 登入完成")
        
        # 步驟3: 導覽到缺勤記錄
        navigate_to_attendance(driver)
        print("✅ 已進入缺勤記錄頁面")
        
        # 步驟4: 解析缺勤數據
        attendance_records = parse_attendance_data(driver)
        
        # 步驟5: 清理和輸出數據
        attendance_df = clean_attendance_data(attendance_records)
        
        if not attendance_df.empty:
            # 嘗試多個輸出位置
            output_attempts = [
                ("attendance_records.csv", "attendance_records.json"),
                (os.path.expanduser("~/Desktop/attendance_records.csv"), os.path.expanduser("~/Desktop/attendance_records.json")),
                (os.path.expanduser("~/Downloads/attendance_records.csv"), os.path.expanduser("~/Downloads/attendance_records.json")),
                (f"attendance_records_{int(time.time())}.csv", f"attendance_records_{int(time.time())}.json")
            ]
            
            csv_saved = False
            json_saved = False
            
            for csv_path, json_path in output_attempts:
                try:
                    # 嘗試寫入CSV
                    if not csv_saved:
                        attendance_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
                        print(f"✅ CSV檔案已保存：{csv_path}")
                        csv_saved = True
                    
                    # 嘗試寫入JSON
                    if not json_saved:
                        attendance_df.to_json(json_path, orient="records", force_ascii=False, indent=2)
                        print(f"✅ JSON檔案已保存：{json_path}")
                        json_saved = True
                    
                    if csv_saved and json_saved:
                        break
                        
                except PermissionError as e:
                    print(f"⚠️ 無法寫入 {csv_path}: 權限不足")
                    continue
                except Exception as e:
                    print(f"⚠️ 寫入失敗 {csv_path}: {e}")
                    continue
            
            if not csv_saved or not json_saved:
                print("❌ 所有輸出位置都失敗，嘗試顯示資料內容：")
                print("\n" + "="*80)
                print("📋 缺勤記錄資料：")
                print("="*80)
                print(attendance_df.to_string(index=False))
                print("="*80)
                
                # 嘗試保存為文字檔
                try:
                    txt_path = f"attendance_records_{int(time.time())}.txt"
                    with open(txt_path, 'w', encoding='utf-8') as f:
                        f.write("缺勤記錄資料\n")
                        f.write("="*50 + "\n")
                        f.write(attendance_df.to_string(index=False))
                    print(f"✅ 已保存為文字檔：{txt_path}")
                except Exception as e:
                    print(f"⚠️ 連文字檔也無法保存：{e}")
            
            print("\n" + "=" * 60)
            print(f"✅ 成功解析缺勤記錄：{len(attendance_df)} 筆")
            
            # 顯示統計資訊
            if '學年' in attendance_df.columns:
                year_counts = attendance_df['學年'].value_counts().sort_index()
                print(f"📊 學年分佈：{dict(year_counts)}")
            
            # 顯示前幾筆資料
            print("\n📋 資料預覽：")
            print(attendance_df.head(10).to_string(index=False))
            
        else:
            print("⚠️ 沒有找到缺勤記錄資料")
            print("💡 請檢查：")
            print("   1. 帳號是否有缺勤記錄")
            print("   2. 頁面結構是否有變化")
            print("   3. 選擇器是否需要更新")
        
        print("\n✅ 爬蟲執行完成！")
        
    except Exception as e:
        print(f"\n❌ 執行失敗: {e}")
        print("🔍 請檢查以下檔案進行除錯：")
        print("   - attendance_debug.html (頁面HTML)")
        print("   - attendance_page_text.txt (頁面文字)")
        print("   - error_attendance.png (錯誤截圖)")
        
        driver.save_screenshot("error_attendance.png")
        save_html(driver, "error_attendance.html")
        raise
        
    finally:
        if not HEADLESS:
            time.sleep(2)
        driver.quit()

if __name__ == "__main__":
    main()