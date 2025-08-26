import logging
import os
from flask import Flask, request, jsonify
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import pymysql # 新增資料庫連線函式庫

app = Flask(__name__)

# 設定 log 方便 debug
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s:%(message)s')
def save_user_to_db(line_user_id, student_id):
    db_config = {
        'host': '125.229.195.9',
        'user': 'login',
        'password': 'Applepie1512',
        'db': 'login_system',
        'charset': 'utf8mb4',
        'cursorclass': pymysql.cursors.DictCursor
    }
    try:
        connection = pymysql.connect(**db_config)
        with connection.cursor() as cursor:
            # 假設資料表叫 users，欄位為 line_user_id 與 student_id
            sql = """
                INSERT INTO users (line_user_id, student_id) 
                VALUES (%s, %s)
                ON DUPLICATE KEY UPDATE student_id=VALUES(student_id)
            """
            cursor.execute(sql, (line_user_id, student_id))
            connection.commit()  # 必須 commit 才會寫入資料庫
        return True
    except pymysql.err.OperationalError as e:
        print(f"資料庫連線錯誤：{e}")
        return False
    except pymysql.err.IntegrityError as e:
        print(f"資料完整性錯誤：{e}")
        return False
    except Exception as e:
        print(f"資料庫操作失敗：{e}")
        return False
    finally:
        if 'connection' in locals() and connection.open:
            connection.close()


# 強制使用 headless 模式，避免 server-side 啟動 GUI 無法顯示
def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument('--no-sandbox')
   ## options.add_argument('--disable-dev-shm-usage')
    ##options.binary_location = '/usr/bin/chromium-browser'
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_window_size(1200, 900)
    driver.implicitly_wait(3)
    return driver

def login_verification(driver, username, password):
    start_url = "https://www.shu.edu.tw/System-info.aspx"
    wait = WebDriverWait(driver, 20)         # 20秒逾時，部分教育網頁需要久一點

    try:
        driver.get(start_url)
        logging.info("Navigate to start page")
        driver.save_screenshot("01_start.png")

        # 點「學生教務系統」連結
        # 點擊連結，並切換視窗
        academic_system_link = wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "學生教務系統"))
        )
        academic_system_link.click()

        # 等待新視窗跳出，切換
        wait.until(lambda d: len(d.window_handles) > 1)
        driver.switch_to.window(driver.window_handles[-1])


        # 確認 SSO 登入頁已載入
        wait.until(EC.url_contains("ap.shu.edu.tw/SSO/login.aspx"))
        driver.save_screenshot("02_sso.png")

        # 填帳號密碼
        username_field = wait.until(EC.visibility_of_element_located((By.ID, 'txtMyId')))
        username_field.clear()
        username_field.send_keys(username)

        password_field = wait.until(EC.visibility_of_element_located((By.ID, 'txtMyPd')))
        password_field.clear()
        password_field.send_keys(password)
        driver.save_screenshot("03_filled.png")
        # 登入
        login_button = wait.until(EC.element_to_be_clickable((By.ID, "btnLogin")))
        driver.execute_script("arguments[0].scrollIntoView();", login_button)  # 保險起見
        try:
            login_button.click()
        except Exception:
            driver.execute_script("arguments[0].click();", login_button)
        logging.info("Clicked login button")

        driver.implicitly_wait(3)
        try:
            error_element = driver.find_element(By.ID, "lblMessage")
            if error_element.is_displayed() and "錯誤" in error_element.text:
                return False, "登入帳號或密碼錯誤！"
        except Exception:
            pass  # 沒有錯誤訊息就略過

        # 如果順利跳轉到首頁
        wait.until(EC.url_contains("ap4.shu.edu.tw/STU1/Index.aspx"))
        return True, "登入成功！"

    except TimeoutException:
        return False, "登入逾時或帳號密碼錯誤！"
    except Exception as e:
        return False, f"登入失敗：{e}"
    except Exception as e:
        driver.save_screenshot("99_error_unknown.png")
        msg = f"其他預期外錯誤 {type(e).__name__} - {str(e)}"
        logging.critical(msg)
        return False, msg

@app.route('/login', methods=['POST'])
def handle_login_request():
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    if not username or not password:
        return jsonify({"success": False, "message": "Missing required parameters"}), 400

    driver = setup_driver()
    try:
        success, message = login_verification(driver, username, password)
    finally:
        driver.quit()
        if success:
            line_user_id = data.get('line_user_id')
            db_success = save_user_to_db(line_user_id, username)
            if db_success:
                return jsonify({"success": True, "message": "Login and binding successful"})
            else:
                return jsonify({"success": False, "message": "Login successful, but database save failed"})
    return jsonify({"success": success, "message": message})

@app.route('/')
def home():
    return "Flask API is running"


if __name__ == '__main__':
    app.run()

