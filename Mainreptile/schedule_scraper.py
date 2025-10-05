# -*- coding: utf-8 -*-
"""
世新大學 SC0106 只抓《個人課表清單一》+ 截圖《個人課表清單二》
- 僅解析 id=GRD_DataGrid 的表格，欄位與網站相同順序
- 不讀、不截、不處理清單二（Schedule1）
- 匯出：timetable_list1.csv / .json / .xlsx
- 新增：截圖保存課表清單二區域
- 若解析不到，會輸出 list1_debug.html 供排查
"""

import os
import time
import re
import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

HEADLESS = False     # 需要背景跑可改 True
HOME_URL = "https://www.shu.edu.tw/"
MAX_WAIT = 25

# 讀 .env 帳密
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

USERNAME = os.getenv("SHU_USERNAME")
PASSWORD = os.getenv("SHU_PASSWORD")
if not USERNAME or not PASSWORD:
    raise SystemExit("❌ 請在 .env 內設定 SHU_USERNAME / SHU_PASSWORD")

LIST1_ORDER = [
    "選別", "課程簡碼", "課程名稱(教材下載)", "開課系級", "學分", "年別",
    "授課老師", "星期節次週別", "教室", "座位序號(行-列)", "備註"
]

def build_driver():
    opt = webdriver.ChromeOptions()
    if HEADLESS:
        opt.add_argument("--headless=new")
    opt.add_argument("--no-sandbox")
    opt.add_argument("--disable-gpu")
    opt.add_argument("--lang=zh-TW")
    # 給大一點的視窗避免欄位自動換行造成解析偏差
    opt.add_argument("--window-size=1600,1400")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opt)

def js_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('click',{bubbles:true}))", el)

def wait_present(driver, by, sel, timeout=MAX_WAIT):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, sel)))

def save_html(driver, path):
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)

def detect_login_error_and_abort(driver):
    """若頁面顯示帳密錯誤等訊息，立刻截圖並結束程式（exit code 2）。"""
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        body_text = ""
    text_low = (body_text or "").lower()
    keywords = [
        "登入帳號或密碼錯誤", "輸入帳號或密碼錯誤", "帳號或密碼錯誤",
        "login failed", "invalid password", "authentication failed"
    ]
    if any(k.lower() in text_low for k in keywords):
        try:
            driver.save_screenshot("login_error.png")
            save_html(driver, "login_error.html")
        except Exception:
            pass
        # 讓父程式能辨識為登入錯誤
        print("❌ 登入失敗：帳號或密碼錯誤", flush=True)
        try:
            driver.quit()
        except Exception:
            pass
        import os
        os._exit(2)

def wait_login_result_or_error(driver, timeout_seconds: int = 8):
    """提交後短暫輪詢：若出現錯誤訊息立即中止；否則返回繼續流程。"""
    end = time.time() + max(1, timeout_seconds)
    last_err = None
    while time.time() < end:
        try:
            # 先檢查常見訊息容器
            try:
                msg = driver.find_element(By.ID, "lblMessage").text
            except Exception:
                msg = ""
            if msg:
                low = msg.lower()
                if any(k in low for k in [
                    '登入帳號或密碼錯誤', '輸入帳號或密碼錯誤', '帳號或密碼錯誤',
                    'login failed', 'invalid password', 'authentication failed']):
                    print("❌ 登入失敗：", msg)
                    try:
                        driver.save_screenshot('login_error.png')
                        save_html(driver, 'login_error.html')
                    except Exception:
                        pass
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    import os
                    os._exit(2)

            # 泛化檢查
            detect_login_error_and_abort(driver)
        except SystemExit:
            raise
        except Exception as e:
            last_err = e
        time.sleep(0.5)
    # 沒檢出錯誤就返回繼續
    return

def text_clean(s: str) -> str:
    s = (s or "").replace("\xa0", " ")
    s = re.sub(r"[ \t\r]+", " ", s)
    s = re.sub(r"\n{2,}", "\n", s).strip()
    return s

# ── 導覽 ─────────────────────────────────────────────────────────────────────
def goto_student_system_from_home(driver):
    driver.get(HOME_URL)
    # 校務系統
    for by, sel in [
        (By.CSS_SELECTOR, "body > div.logosearch-area > div.n2021-area > p > a:nth-child(4)"),
        (By.XPATH, "//a[contains(@href,'System-info.aspx')]"),
        (By.XPATH, "//a[contains(.,'校務系統')]"),
    ]:
        try:
            el = driver.find_element(by, sel)
            js_click(driver, el)
            break
        except Exception:
            continue
    else:
        save_html(driver, "fail_sys_link.html")
        driver.save_screenshot("fail_sys_link.png")
        raise RuntimeError("找不到『校務系統』連結")

    time.sleep(0.6)
    # 學生教務系統
    for by, sel in [
        (By.CSS_SELECTOR, "body > div:nth-child(10) > div > div.sm-page-all-area > div:nth-child(2) > div.ct-sub-sbox.ct-sub-nsbox.ct-sub-nsortbox > a:nth-child(9)"),
        (By.XPATH, "//a[contains(@href,'stulb.shu.edu.tw')]"),
        (By.XPATH, "//a[contains(.,'學生教務系統')]"),
    ]:
        try:
            el = driver.find_element(by, sel)
            js_click(driver, el)
            break
        except Exception:
            continue
    else:
        driver.execute_script("window.open('https://stulb.shu.edu.tw/','_blank');")

    driver.switch_to.window(driver.window_handles[-1])

def login_if_needed(driver):
    try:
        u = WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'],input[autocomplete='username']"))
        )
        p = wait_present(driver, By.CSS_SELECTOR, "input[type='password'],input[autocomplete='current-password']", 8)
        u.clear(); u.send_keys(USERNAME)
        p.clear(); p.send_keys(PASSWORD)
        try:
            btn = driver.find_element(By.CSS_SELECTOR, "input[type='submit'],button[type='submit']")
            js_click(driver, btn)
        except NoSuchElementException:
            p.submit()
        time.sleep(0.8)
        # 檢查是否顯示登入錯誤（包含輪詢）
        wait_login_result_or_error(driver, timeout_seconds=8)
    except TimeoutException:
        pass  # 沒有登入畫面就 SSO 直通

def open_sc0106(driver):
    driver.switch_to.default_content()
    try:
        driver.switch_to.frame("main")
    except Exception:
        # 若找不到 frame，記錄並改用當前內容繼續，避免整段流程中斷
        try:
            save_html(driver, "frameset_outer.html")
            driver.save_screenshot("no_main_frame.png")
        except Exception:
            pass
        print("[WARN] 找不到 main frame，改用目前頁面繼續。")

    wait_present(driver, By.CSS_SELECTOR, ".label", 15)

    # 課務作業
    js_click(driver, driver.find_element(By.XPATH, "//span[@class='label' and contains(.,'課務作業')]"))
    time.sleep(0.2)

    # SC0106
    for by, sel in [
        (By.XPATH, "//span[contains(.,'SC0106-學生課表查詢')]"),
        (By.XPATH, "//span[contains(.,'學生課表查詢')]"),
    ]:
        try:
            js_click(driver, driver.find_element(by, sel))
            break
        except Exception:
            continue
    else:
        driver.save_screenshot("click_sc0106_fail.png")
        raise RuntimeError("點不到 SC0106-學生課表查詢")

    time.sleep(1.0)

def select_latest_and_search(driver):
    driver.switch_to.default_content()
    driver.switch_to.frame("main")
    # 學年/學期設為最大
    driver.execute_script("""
        const isYear = t => /^\\d{3}$/.test((t||'').trim());
        const isTerm = t => /^(1|2|3|4)$/.test((t||'').trim());
        const sels = Array.from(document.querySelectorAll('select'));
        let ySel=null,tSel=null, yMax=-1, tMax=-1;
        for (const s of sels) {
          const opts = Array.from(s.options).map(o => (o.textContent||o.value||'').trim());
          const ys = opts.filter(isYear).map(Number);
          if (ys.length) { const m=Math.max(...ys); if (m>=yMax){yMax=m;ySel=s;} }
          const ts = opts.filter(isTerm).map(v=>Number(String(v).replace(/\\D/g,'')));
          if (ts.length) { const m=Math.max(...ts); if (m>=tMax){tMax=m;tSel=s;} }
        }
        function setTo(sel, v){
          if(!sel) return false;
          const i = Array.from(sel.options).findIndex(o => (o.textContent||o.value||'').trim()==String(v));
          if(i>=0){ sel.selectedIndex=i; sel.dispatchEvent(new Event('change',{bubbles:true})); return true; }
          return false;
        }
        setTo(ySel, yMax);
        setTo(tSel, tMax);
    """)
    # 搜尋
    for by, sel in [(By.ID, "SRH_search_button"),
                    (By.XPATH, "//input[@type='submit' and contains(@value,'搜尋')]")]:
        try:
            js_click(driver, driver.find_element(by, sel))
            break
        except Exception:
            continue

    # 只等清單一的資料表
    wait_present(driver, By.ID, "GRD_DataGrid", 20)
    time.sleep(0.3)

# ── 只解析清單一 ─────────────────────────────────────────────────────────────
def parse_list1(driver) -> pd.DataFrame:
    driver.switch_to.default_content()
    driver.switch_to.frame("main")

    # 用 JS 從 #GRD_DataGrid 精準把 header + 每列文字抽出；避免拿到任何巢狀表格
    data = driver.execute_script("""
        const tbl = document.getElementById('GRD_DataGrid');
        if(!tbl) return {headers:[], rows:[]};

        const clean = (s) => (s||'').replace(/\\u00a0/g,' ').replace(/[ \\t\\r]+/g,' ')
                                    .replace(/\\n{2,}/g,'\\n').trim();

        const ths = Array.from(tbl.querySelectorAll(':scope > tbody > tr:first-child > td, :scope > tbody > tr:first-child > th'));
        const headers = ths.map(c => clean(c.innerText));

        const trs = Array.from(tbl.querySelectorAll(':scope > tbody > tr')).slice(1);
        const rows = [];
        for (const tr of trs) {
          const tds = Array.from(tr.querySelectorAll(':scope > td, :scope > th'));
          if (!tds.length) continue;
          const vals = [];
          for (let i=0;i<tds.length;i++){
            const td = tds[i];
            let txt = '';
            if (headers[i] && headers[i].includes('課程名稱')) {
              const a = td.querySelector('a');
              if (a) txt = a.innerText;
            }
            if (!txt) txt = td.innerText;
            vals.push(clean(txt));
          }
          // 濾掉整列空白
          if (vals.some(v => v && v.length)) rows.push(vals);
        }
        return {headers, rows};
    """)

    headers = [h for h in data.get("headers", [])]
    rows = data.get("rows", [])

    if not headers or not rows:
        # 萬一 header 沒抓到，用固定欄序備援
        headers = LIST1_ORDER[:]

    # 有些列會比表頭多/少，這裡對齊一下長度
    fixed_rows = []
    for r in rows:
        if len(r) < len(headers):
            r = r + [""] * (len(headers) - len(r))
        elif len(r) > len(headers):
            r = r[:len(headers)]
        fixed_rows.append(r)

    df = pd.DataFrame(fixed_rows, columns=headers)

    # 以網站欄序輸出（缺的就忽略，多的放最後）
    keep = [h for h in LIST1_ORDER if h in df.columns]
    others = [c for c in df.columns if c not in keep]
    return df[keep + others]

# ── 新增：截圖清單二區域 ───────────────────────────────────────────────────────
def screenshot_list2(driver):
    """截圖課表清單二區域並保存"""
    try:
        driver.switch_to.default_content()
        driver.switch_to.frame("main")
        
        # 先保存完整頁面用於調試
        driver.save_screenshot("debug_full_page.png")
        print("🔍 調試：完整頁面截圖已保存 debug_full_page.png")
        
        # 使用JavaScript來精確找到課表清單二
        list2_info = driver.execute_script("""
            // 尋找課表清單二的多種策略
            let targetTable = null;
            let strategy = '';
            
            // 策略1: 尋找包含星期的表格（非GRD_DataGrid）
            const tables = Array.from(document.querySelectorAll('table'));
            for (let table of tables) {
                const tableId = table.id || '';
                if (tableId === 'GRD_DataGrid') continue; // 跳過清單一
                
                const text = table.innerText || '';
                const hasWeekdays = ['星期一', '星期二', '星期三', '星期四', '星期五'].some(day => text.includes(day));
                if (hasWeekdays) {
                    targetTable = table;
                    strategy = 'weekday_table';
                    break;
                }
            }
            
            // 策略2: 尋找藍色背景的表格
            if (!targetTable) {
                for (let table of tables) {
                    if (table.id === 'GRD_DataGrid') continue;
                    
                    const cells = Array.from(table.querySelectorAll('td'));
                    const hasBlueBackground = cells.some(cell => {
                        const bgColor = window.getComputedStyle(cell).backgroundColor;
                        const bgcolor = cell.getAttribute('bgcolor');
                        return bgColor.includes('rgb') && bgColor !== 'rgba(0, 0, 0, 0)' || 
                               (bgcolor && bgcolor !== '');
                    });
                    
                    if (hasBlueBackground) {
                        targetTable = table;
                        strategy = 'blue_background';
                        break;
                    }
                }
            }
            
            // 策略3: 找GRD_DataGrid後面的第一個表格
            if (!targetTable) {
                const dataGrid = document.getElementById('GRD_DataGrid');
                if (dataGrid) {
                    let nextEl = dataGrid.nextElementSibling;
                    while (nextEl) {
                        if (nextEl.tagName === 'TABLE') {
                            targetTable = nextEl;
                            strategy = 'next_after_datagrid';
                            break;
                        }
                        const table = nextEl.querySelector('table');
                        if (table) {
                            targetTable = table;
                            strategy = 'nested_after_datagrid';
                            break;
                        }
                        nextEl = nextEl.nextElementSibling;
                    }
                }
            }
            
            // 策略4: 尋找最下方的大表格
            if (!targetTable) {
                const allTables = Array.from(document.querySelectorAll('table'));
                const largeTables = allTables.filter(t => {
                    const rect = t.getBoundingClientRect();
                    return rect.width > 500 && rect.height > 200 && t.id !== 'GRD_DataGrid';
                });
                if (largeTables.length > 0) {
                    targetTable = largeTables[largeTables.length - 1]; // 取最後一個大表格
                    strategy = 'large_table';
                }
            }
            
            if (targetTable) {
                const rect = targetTable.getBoundingClientRect();
                return {
                    found: true,
                    strategy: strategy,
                    element: targetTable,
                    x: rect.x + window.pageXOffset,
                    y: rect.y + window.pageYOffset,
                    width: rect.width,
                    height: rect.height,
                    text_preview: targetTable.innerText.substring(0, 100)
                };
            } else {
                return { found: false };
            }
        """)
        
        if not list2_info.get('found'):
            print("⚠️ 無法找到課表清單二，使用頁面下半部截圖")
            # 截圖頁面下半部
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight * 0.5);")
            time.sleep(0.5)
            driver.save_screenshot("timetable_list2_bottom_half.png")
            return
        
        print(f"✅ 找到清單二，使用策略：{list2_info['strategy']}")
        print(f"📝 內容預覽：{list2_info['text_preview']}...")
        
        # 滾動到清單二位置
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", list2_info['element'])
        time.sleep(0.8)
        
        # 截圖整個頁面
        driver.save_screenshot("temp_full_screenshot.png")
        
        # 使用PIL裁切出清單二區域
        try:
            from PIL import Image
            
            # 打開完整截圖
            full_image = Image.open("temp_full_screenshot.png")
            
            # 獲取瀏覽器的縮放比例
            device_pixel_ratio = driver.execute_script("return window.devicePixelRatio || 1;")
            
            # 計算實際像素位置（考慮縮放比例），並加大padding確保包含標題
            padding = 50  # 增加padding確保包含標題
            left = max(0, int(list2_info['x'] * device_pixel_ratio) - padding)
            top = max(0, int(list2_info['y'] * device_pixel_ratio) - padding)
            right = min(full_image.width, int((list2_info['x'] + list2_info['width']) * device_pixel_ratio) + padding)
            bottom = min(full_image.height, int((list2_info['y'] + list2_info['height']) * device_pixel_ratio) + padding)
            
            print(f"🔍 裁切區域：({left}, {top}) 到 ({right}, {bottom})")
            
            # 裁切並保存
            if right > left and bottom > top:
                cropped_image = full_image.crop((left, top, right, bottom))
                cropped_image.save("timetable_list2.png")
                print("📸 課表清單二截圖已保存：timetable_list2.png")
            else:
                print("⚠️ 裁切區域無效，保存完整截圖")
                full_image.save("timetable_list2_full.png")
            
            # 刪除臨時檔案
            os.remove("temp_full_screenshot.png")
            
        except ImportError:
            print("⚠️ 需要安裝 Pillow 套件：pip install Pillow")
            print("📸 完整頁面截圖已保存")
            os.rename("temp_full_screenshot.png", "timetable_list2_fullpage.png")
            
        except Exception as crop_error:
            print(f"⚠️ 裁切時發生錯誤：{crop_error}")
            print("📸 保存完整頁面截圖作為備用")
            if os.path.exists("temp_full_screenshot.png"):
                os.rename("temp_full_screenshot.png", "timetable_list2_backup.png")
            
    except Exception as e:
        print(f"⚠️ 截圖清單二時發生錯誤：{e}")
        try:
            # 最終備用方案：截圖整個頁面
            driver.save_screenshot("timetable_list2_emergency.png")
            print("📸 緊急備用截圖已保存：timetable_list2_emergency.png")
        except Exception:
            pass

# ── 主程式 ───────────────────────────────────────────────────────────────────
def main():
    driver = build_driver()
    try:
        print("🚀 啟動：只抓清單一 + 截圖清單二")
        goto_student_system_from_home(driver)
        login_if_needed(driver)
        open_sc0106(driver)
        select_latest_and_search(driver)

        df = parse_list1(driver)
        if df.empty:
            save_html(driver, "list1_debug.html")
            raise RuntimeError("清單一解析不到資料；已輸出 list1_debug.html 供檢查")

        # 匯出清單一（不做 pivot/merge/展開節次，完全照清單一）
        df.to_csv("timetable_list1.csv", index=False, encoding="utf-8-sig")
        df.to_json("timetable_list1.json", orient="records", force_ascii=False, indent=2)

        with pd.ExcelWriter("timetable_list1.xlsx", engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="清單一")
            ws = writer.sheets["清單一"]
            # 簡單寬度（可依需求調整）
            width_map = {
                "選別": 6, "課程簡碼": 18, "課程名稱(教材下載)": 28, "開課系級": 16,
                "學分": 6, "年別": 6, "授課老師": 12, "星期節次週別": 20, "教室": 18,
                "座位序號(行-列)": 14, "備註": 36
            }
            for col, w in width_map.items():
                if col in df.columns:
                    ci = df.columns.get_loc(col)
                    ws.set_column(ci, ci, w)

        print("✅ 清單一完成：timetable_list1.(csv/json/xlsx) 已產生")
        
        # 新增：截圖清單二區域
        print("📸 開始截圖課表清單二...")
        screenshot_list2(driver)
        
        print("🎉 全部完成：清單一資料 + 清單二截圖")

    finally:
        try:
            driver.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()