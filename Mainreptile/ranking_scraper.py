# -*- coding: utf-8 -*-
"""
世新大學 歷年名次爬蟲
- 從世新校網進入學生教務系統，爬取個人歷年名次記錄
- 修正：學分欄名 (避免寫成名分)
- 修正：名次／人數在 CSV 被 Excel 誤判成日期
  - CSV：名次／人數以 ="..." 文字包裹，且斜線改全形「／」
  - XLSX：另存一份，名次／人數兩欄整欄鎖為文字格式
- 加值：拆出名次_班/組/系 與 人數_班/組/系 數字欄位
"""

import time
import os
from typing import List, Tuple
import re
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# ───── 載入環境變數 (.env 需有 SHU_USERNAME / SHU_PASSWORD) ─────
try:
    from dotenv import load_dotenv
    load_dotenv()
    print("✅ 已載入 .env 檔案")
except ImportError:
    print("⚠️ 未安裝 python-dotenv，請執行: pip install python-dotenv")

# ========= 設定區 =========
USERNAME = os.getenv('SHU_USERNAME')
PASSWORD = os.getenv('SHU_PASSWORD')

if not USERNAME or not PASSWORD:
    print("❌ 錯誤：請在 .env 檔案中設定 SHU_USERNAME 和 SHU_PASSWORD")
    exit(1)

HOME_URL = "https://www.shu.edu.tw/"
HEADLESS = False
MAX_WAIT = 25

print(f"🔐 使用帳號：{USERNAME[:3]}***{USERNAME[-3:] if len(USERNAME) > 6 else '***'}")

# ---------------- 基礎工具函數 ----------------
def build_driver():
    opt = webdriver.ChromeOptions()
    if HEADLESS:
        opt.add_argument("--headless=new")
    opt.add_argument("--no-sandbox")
    opt.add_argument("--disable-gpu")
    opt.add_argument("--window-size=1440,900")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opt)

def js_click(driver, el):
    driver.execute_script("""
        const el = arguments[0];
        el.scrollIntoView({block:'center'});
        try{ el.dispatchEvent(new MouseEvent('mouseover',{bubbles:true})); }catch(e){}
        try{ el.dispatchEvent(new MouseEvent('mousedown',{bubbles:true})); }catch(e){}
        try{ el.dispatchEvent(new MouseEvent('mouseup',{bubbles:true})); }catch(e){}
        try{ el.click(); }catch(e){}
    """, el)

def find_and_js_click(driver, selector: str, by="css") -> bool:
    try:
        el = driver.find_element(By.CSS_SELECTOR, selector) if by=="css" else driver.find_element(By.XPATH, selector)
        js_click(driver, el)
        return True
    except Exception:
        return False

def click_first_working(driver, selectors: List[Tuple[str, str]]) -> bool:
    for by, sel in selectors:
        if find_and_js_click(driver, sel, by=by):
            return True
    return False

def save_html(driver, path):
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)

def _die(driver, msg, png, html):
    driver.save_screenshot(png)
    save_html(driver, html)
    raise RuntimeError(f"{msg}；已存 {png} / {html}")

# ---------------- 導覽函數 ----------------
def goto_student_system_from_home(driver):
    driver.get(HOME_URL)
    ok = click_first_working(driver, [
        ("css",  "body > div.logosearch-area > div.n2021-area > p > a:nth-child(4)"),
        ("xpath","//a[contains(@href,'System-info.aspx')]"),
    ])
    if not ok:
        _die(driver, "找不到『校務系統』連結", "fail_sys_link.png", "fail_sys_link.html")
    time.sleep(0.6)

    ok = click_first_working(driver, [
        ("css",  "body > div:nth-child(10) .ct-sub-nsortbox > a:nth-child(9)"),
        ("css",  "body > div:nth-child(11) .ct-sub-nsortbox > a:nth-child(9)"),
        ("xpath","//a[contains(@href,'stulb.shu.edu.tw')]"),
    ])
    if not ok:
        driver.execute_script("window.open('https://stulb.shu.edu.tw/','_blank');")

    driver.switch_to.window(driver.window_handles[-1])

def login_if_needed(driver):
    try:
        u = WebDriverWait(driver, 12).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'],input[autocomplete='username']"))
        )
        p = WebDriverWait(driver, 12).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password'],input[autocomplete='current-password']"))
        )
        u.clear(); u.send_keys(USERNAME)
        p.clear(); p.send_keys(PASSWORD)
        try:
            btn = driver.find_element(By.CSS_SELECTOR, "input[type='submit'],button[type='submit']")
            js_click(driver, btn)
        except NoSuchElementException:
            p.submit()
        time.sleep(0.8)
        # 提交後短暫輪詢錯誤訊息，若偵測到立即結束
        end = time.time() + 8
        while time.time() < end:
            try:
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
        pass

def open_ranking_page(driver):
    driver.switch_to.default_content()
    try:
        driver.switch_to.frame("main")
    except Exception:
        _die(driver, "找不到 main frame", "no_main_frame.png", "frameset_outer.html")

    WebDriverWait(driver, 20).until(lambda d: d.execute_script("return document.readyState") == "complete")
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".label")))
    except TimeoutException:
        _die(driver, "main frame 內未出現 .label", "main_no_labels.png", "main_no_labels.html")

    def click_label(text) -> bool:
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

    if not click_label("成績作業"):
        ok = click_first_working(driver, [
            ("xpath", "//span[@class='label' and normalize-space()='成績作業']"),
            ("css", "#app > div > ul > div > div:nth-child(3) > span"),
        ])
        if not ok:
            _die(driver, "點不到『成績作業』", "click_fail_grade.png", "click_fail_grade.html")
    time.sleep(0.6)

    ranking_keywords = [
        "SD0104-歷年(學期)名次查詢",
        "SD0104",
        "歷年(學期)名次查詢",
        "歷年名次",
        "名次查詢"
    ]
    success = False
    for keyword in ranking_keywords:
        if click_label(keyword):
            print(f"成功點擊: {keyword}")
            success = True
            break

    if not success:
        ok = click_first_working(driver, [
            ("css", "#app > div > ul > div > div:nth-child(3) > div > div.bar-menu-items > div:nth-child(2) > span"),
            ("xpath", "//span[contains(text(), 'SD0104')]"),
            ("xpath", "//span[contains(text(), '歷年') and contains(text(), '名次')]"),
        ])
        if ok:
            success = True

    if not success:
        success = driver.execute_script("""
            const keywords = ["SD0104", "歷年", "名次"];
            const els = Array.from(document.querySelectorAll('.label, a, button, span, div'));
            for (const el of els) {
              const t = (el.textContent||'').replace(/\s+/g,'');
              for (const keyword of keywords) {
                if (t.includes(keyword)) {
                  el.scrollIntoView({block:'center'});
                  try { el.click(); return true; } catch(e){}
                  try { el.parentElement?.click(); return true; } catch(e){}
                }
              }
            }
            return false;
        """)

    if not success:
        _die(driver, "點不到『SD0104-歷年(學期)名次查詢』選單", "click_fail_ranking.png", "click_fail_ranking.html")

    time.sleep(3.0)

# ---------------- 解析 + 清理函數 ----------------
def parse_ranking_data(driver) -> pd.DataFrame:
    """解析歷年名次數據"""
    driver.switch_to.default_content()
    driver.switch_to.frame("main")

    time.sleep(2)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(1)

    print("🔍 開始解析歷年名次...")
    save_html(driver, "ranking_debug.html")

    records = []

    try:
        tables = driver.find_elements(By.TAG_NAME, "table")
        for i, table in enumerate(tables):
            table_text = table.text
            if any(keyword in table_text for keyword in ['學年度', '學期', '平均', '名次', 'SD0104']):
                rows = table.find_elements(By.TAG_NAME, "tr")
                headers = []
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if not cells:
                        cells = row.find_elements(By.TAG_NAME, "th")
                    if not cells:
                        continue
                    cell_texts = [cell.text.strip() for cell in cells]

                    # 找表頭
                    if not headers and any('學年度' in t or '學期' in t or '平均' in t for t in cell_texts):
                        headers = cell_texts
                        continue

                    # 資料列
                    if len(cell_texts) >= 6 and any(cell_texts):
                        try:
                            rec = {
                                '學年度': cell_texts[0],
                                '學期':   cell_texts[1],
                                '學分':   cell_texts[2],              # ✅ 修正：學分（不是名分）
                                '平均':   cell_texts[3],
                                '名次':   cell_texts[4],
                                '人數':   cell_texts[5],
                            }

                            # 基本型別檢查
                            if (rec['學年度'].isdigit() and
                                rec['學期'] in ['1', '2'] and
                                re.match(r'^\d+(\.\d+)?$', rec['平均'])):
                                records.append(rec)
                        except Exception as e:
                            print(f"解析行失敗: {e}, 內容: {cell_texts}")
                            continue

        # 後備：若表格沒抓到，解析底部統計（保留原邏輯）
        if not records:
            print("🔄 嘗試解析底部統計資訊...")
            page_text = driver.find_element(By.TAG_NAME, "body").text
            lines = [line.strip() for line in page_text.split('\n') if line.strip()]
            for line in lines:
                m = re.search(r'該生至\s*(\d+)\s*學年第\s*(\d+)\s*學期止.*?(\d+)\s*/\s*(\d+)', line)
                if m:
                    year, semester, rank, total = m.groups()
                    records.append({
                        '學年度': year, '學期': semester,
                        '排名類型': '累計排名',
                        '名次': f"{rank}／{total}",  # 用全形斜線
                        '備註': line
                    })

    except Exception as e:
        print(f"❌ 解析過程發生錯誤: {e}")
        try:
            page_text = driver.find_element(By.TAG_NAME, "body").text
            with open("ranking_page_text_debug.txt", "w", encoding="utf-8") as f:
                f.write("頁面完整文字內容:\n" + "="*50 + "\n" + page_text)
            print("📝 已保存頁面文字到 ranking_page_text_debug.txt")
        except:
            pass

    print(f"✅ 解析完成：{len(records)} 筆名次記錄")
    df = pd.DataFrame(records) if records else pd.DataFrame()
    if not df.empty:
        df = clean_ranking_df(df)
    return df

def _to_fullwidth_slash(s: str) -> str:
    """把 'a / b / c' 轉成 'a／b／c'，避免 Excel 自動變日期"""
    if s is None:
        return s
    return re.sub(r'\s*/\s*', '／', str(s))

def _split_triplet_to_cols(series: pd.Series, prefix: str) -> pd.DataFrame:
    """將 'x／y／z' 拆成三欄（保留數字），不存在則回 None"""
    def _split(s):
        nums = re.findall(r'\d+', str(s) if s is not None else '')
        nums = nums[:3] + [None] * (3 - len(nums))
        return [int(n) if n is not None else None for n in nums]
    return series.apply(_split).apply(pd.Series).set_axis([f'{prefix}_班', f'{prefix}_組', f'{prefix}_系'], axis=1)

def clean_ranking_df(df: pd.DataFrame) -> pd.DataFrame:
    """清理名次資料"""
    if df.empty:
        return df

    # 欄名一致化
    if '學年度' in df.columns:
        df = df.rename(columns={'學年度': '學年'})

    # 型別修正
    if '學年' in df.columns:
        df['學年'] = df['學年'].astype(str)
    if '學期' in df.columns:
        df['學期'] = df['學期'].astype(str)
    if '平均' in df.columns:
        df['平均'] = pd.to_numeric(df['平均'], errors='coerce')

    # ✅ 防止 Excel 把名次/人數當日期：改用全形斜線，並另外拆成數字欄
    if '名次' in df.columns:
        df['名次'] = df['名次'].map(_to_fullwidth_slash)
        rank_cols = _split_triplet_to_cols(df['名次'], '名次')
        df = pd.concat([df, rank_cols], axis=1)

    if '人數' in df.columns:
        df['人數'] = df['人數'].map(_to_fullwidth_slash)
        count_cols = _split_triplet_to_cols(df['人數'], '人數')
        df = pd.concat([df, count_cols], axis=1)

    # 排序（若欄位存在）
    sort_cols = [c for c in ['學年', '學期'] if c in df.columns]
    if sort_cols:
        df = df.sort_values(sort_cols).reset_index(drop=True)

    # 欄位順序（直覺）
    ordered = [c for c in ['學年', '學期', '學分', '平均', '名次', '人數',
                           '名次_班', '名次_組', '名次_系',
                           '人數_班', '人數_組', '人數_系'] if c in df.columns]
    df = df[ordered + [c for c in df.columns if c not in ordered]]
    return df

# ---------------- 主程式 ----------------
def main():
    driver = build_driver()
    try:
        print("🚀 開始執行歷年名次爬蟲...")
        goto_student_system_from_home(driver)
        print("✅ 已進入學生教務系統")

        login_if_needed(driver)
        print("✅ 登入完成")

        open_ranking_page(driver)
        print("✅ 已開啟歷年名次頁面")

        ranking_df = parse_ranking_data(driver)

        # --- 輸出檔案（CSV 防日期 + 另存 XLSX 鎖文字） ---
        if not ranking_df.empty:
            # 給 CSV 用的副本：名次/人數包成 ="..."，並保留全形斜線
            df_csv = ranking_df.copy()
            for col in ['名次', '人數']:
                if col in df_csv.columns:
                    df_csv[col] = df_csv[col].astype(str).str.replace('/', '／', regex=False)
                    df_csv[col] = '="' + df_csv[col].str.replace('"', '""') + '"'

            output_attempts = [
                ("ranking_records.csv", "ranking_records.json", "ranking_records.xlsx"),
                (os.path.expanduser("~/Desktop/ranking_records.csv"),
                 os.path.expanduser("~/Desktop/ranking_records.json"),
                 os.path.expanduser("~/Desktop/ranking_records.xlsx")),
                (os.path.expanduser("~/Downloads/ranking_records.csv"),
                 os.path.expanduser("~/Downloads/ranking_records.json"),
                 os.path.expanduser("~/Downloads/ranking_records.xlsx")),
                (f"ranking_records_{int(time.time())}.csv",
                 f"ranking_records_{int(time.time())}.json",
                 f"ranking_records_{int(time.time())}.xlsx"),
            ]

            csv_saved = json_saved = xlsx_saved = False
            for csv_path, json_path, xlsx_path in output_attempts:
                try:
                    if not csv_saved:
                        df_csv.to_csv(csv_path, index=False, encoding="utf-8-sig")
                        print(f"✅ CSV檔案已保存：{csv_path}")
                        csv_saved = True

                    if not json_saved:
                        ranking_df.to_json(json_path, orient="records", force_ascii=False, indent=2)
                        print(f"✅ JSON檔案已保存：{json_path}")
                        json_saved = True

                    if not xlsx_saved:
                        # 另存 XLSX，指定名次/人數整欄為文字格式
                        with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
                            ranking_df.to_excel(writer, index=False, sheet_name="歷年名次")
                            wb = writer.book
                            ws = writer.sheets["歷年名次"]
                            text_fmt = wb.add_format({'num_format': '@'})  # 文字格式
                            for col_name in ['名次', '人數']:
                                if col_name in ranking_df.columns:
                                    ci = ranking_df.columns.get_loc(col_name)
                                    ws.set_column(ci, ci, 16, text_fmt)  # 欄寬 16，整欄套文字
                        print(f"✅ XLSX檔案已保存（名次/人數鎖文字）：{xlsx_path}")
                        xlsx_saved = True

                    if csv_saved and json_saved and xlsx_saved:
                        break

                except PermissionError:
                    print(f"⚠️ 無法寫入 {csv_path}: 權限不足"); continue
                except Exception as e:
                    print(f"⚠️ 寫入失敗：{e}"); continue

            if not (csv_saved and json_saved and xlsx_saved):
                print("❌ 部分輸出失敗，以下為資料預覽：")
                print("\n" + "="*60)
                print(ranking_df.to_string(index=False))
                print("="*60)

            print(f"✅ 已輸出名次資料：{len(ranking_df)} 筆")

            if '學年' in ranking_df.columns:
                year_counts = ranking_df['學年'].value_counts().sort_index()
                print(f"   學年分佈：{dict(year_counts)}")

            if not ranking_df.empty:
                print("\n📊 資料預覽：")
                print(ranking_df.to_string(index=False, max_colwidth=15))
        else:
            print("⚠️ 沒有找到名次資料")

        print("\n✅ 爬蟲執行完成！")

    except Exception as e:
        print(f"❌ 執行失敗: {e}")
        driver.save_screenshot("error_ranking.png")
        save_html(driver, "error_ranking.html")
        try:
            page_text = driver.find_element(By.TAG_NAME, "body").text
            with open("ranking_page_text_debug.txt", "w", encoding="utf-8") as f:
                f.write(page_text)
            print("已儲存除錯資訊到 ranking_page_text_debug.txt")
        except:
            pass
        raise
    finally:
        time.sleep(2 if not HEADLESS else 0)
        driver.quit()

if __name__ == "__main__":
    main()
