# -*- coding: utf-8 -*-
import os
import sys
import time
import glob
import subprocess
from pathlib import Path
from typing import Optional, Dict, Any

import pandas as pd
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from dotenv import load_dotenv
from dotenv import set_key

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("APP_SECRET", "dev-secret")  # for flash()

# === 設定：請把路徑改成你電腦上實際的檔名 ===
# 建議把你的四個爬蟲檔擺一起（或改成絕對路徑）
# 你的實際檔名位於「Main reptile/」資料夾中，對應如下：
# - 課表（清單一）：Main reptile/schedule_scraper.py → 產出 timetable_list1.csv
# - 歷年成績：      Main reptile/grade.py        → 產出 grades_courses_fixed.csv / grades_summary_fixed.csv
# - 歷年名次：      Main reptile/ranking_scraper.py → 產出 ranking_records.csv
# - 出缺勤記錄：    Main reptile/attendance_scraper.py → 產出 attendance_records.csv
BASE_DIR = Path(__file__).parent.resolve()
SCRIPTS = {
    "timetable": str((BASE_DIR / "Mainreptile" / "schedule_scraper.py").resolve()),
    "grades":    str((BASE_DIR / "Mainreptile" / "grade.py").resolve()),
    "ranking":   str((BASE_DIR / "Mainreptile" / "ranking_scraper.py").resolve()),
    "attendance":str((BASE_DIR / "Mainreptile" / "attendance_scraper.py").resolve()),
}

# 各腳本跑完後**預期**會產生的檔案（用來找最新一份）
OUTPUTS = {
    "timetable": ["timetable_list1.csv"],
    "grades":    ["grades_courses_fixed.csv", "grades_summary_fixed.csv"],
    "ranking":   ["ranking_records.csv"],
    "attendance":["attendance_records.csv"],
}

# CSV 顯示時的預設欄位（有就秀；沒有就自動顯示全部）
DEFAULT_COLUMNS = {
    "timetable": ["選別","課程簡碼","課程名稱(教材下載)","開課系級","學分","授課老師","星期節次週別","教室","備註"],
    "grades":    ["學年","選別","科目","上學期_學分","上學期_成績","下學期_學分","下學期_成績"],
    "ranking":   ["學年","學期","學分","平均","名次","人數"],
    "attendance":["學年","學期","課程代碼","課程名稱","教師","缺勤狀態","曠課次數","扣考時數","備註"],
}

# 在 Windows 通常用 "python"，在虛擬環境/其他系統用 sys.executable 比較穩
PYTHON_BIN = sys.executable or "python"

# 使用者資料夾根目錄
DATA_ROOT = Path("data")
DATA_ROOT.mkdir(exist_ok=True)
LAST_USER_FILE = DATA_ROOT/".last_username"

# 各類型預設逾時秒數（避免子行程無限卡住）
SCRIPT_TIMEOUTS = {
    "timetable": int(os.getenv("TIMEOUT_TIMETABLE", "240")),
    "grades":    int(os.getenv("TIMEOUT_GRADES", "300")),
    "ranking":   int(os.getenv("TIMEOUT_RANKING", "300")),
    "attendance":int(os.getenv("TIMEOUT_ATTENDANCE", "300")),
}


def run_script(kind: str, env_override: Dict[str, Any], work_dir: Optional[Path] = None) -> Dict[str, Any]:
    """
    呼叫對應的爬蟲腳本。回傳 process returncode。
    會把 SHU_USERNAME / SHU_PASSWORD / HEADLESS 等環境變數覆寫進去（不落地存檔）。
    """
    script = SCRIPTS.get(kind)
    if not script or not Path(script).exists():
        raise FileNotFoundError(f"找不到爬蟲腳本：{script}（請確認 SCRIPTS 設定與檔名）")

    env = os.environ.copy()
    env.update({k: str(v) for k, v in env_override.items() if v is not None})
    # 強制子行程以 UTF-8 輸出，避免 Windows cp950 解碼錯誤
    env["PYTHONIOENCODING"] = "utf-8"

    # 讓 selenium 在 server 上能跑
    if "HEADLESS" not in env:
        env["HEADLESS"] = os.getenv("HEADLESS", "True")

    # 執行
    run_cwd = Path(work_dir) if work_dir else Path.cwd()
    print(f"[RUN] {PYTHON_BIN} {script} (cwd={run_cwd})")
    try:
        proc = subprocess.run(
            [PYTHON_BIN, script],
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=str(run_cwd),
            timeout=SCRIPT_TIMEOUTS.get(kind, 300)
        )
        ret_code = proc.returncode
        stdout = proc.stdout or ""
        stderr = proc.stderr or ""
    except subprocess.TimeoutExpired as te:
        ret_code = 124  # 常見的 timeout 代碼
        stdout = (te.stdout or "") if hasattr(te, "stdout") else ""
        stderr = (te.stderr or "") if hasattr(te, "stderr") else ""
        stderr += "\n[ERROR] 子行程執行逾時，已中止。"
    # 把標準輸出／錯誤留檔方便除錯（存到使用者資料夾下的 logs/）
    ts = int(time.time())
    logs_dir = run_cwd/"logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    out_path = logs_dir/f"{kind}_{ts}.out.txt"
    err_path = logs_dir/f"{kind}_{ts}.err.txt"
    out_path.write_text(stdout, encoding="utf-8")
    err_path.write_text(stderr, encoding="utf-8")
    print(f"[RET] code={ret_code}")
    return {"code": ret_code, "out": str(out_path), "err": str(err_path)}


def _log_contains_login_error(text: str) -> bool:
    """粗略判斷是否為登入帳號/密碼錯誤（關鍵字比對）。"""
    if not text:
        return False
    keywords = [
        "帳號或密碼錯誤", "密碼錯誤", "登入失敗", "Login failed", "invalid password",
        "authentication failed", "驗證失敗", "輸入錯誤", "帳密錯誤"
    ]
    low = text.lower()
    return any(k.lower() in low for k in keywords)


def _diagnose_message(out_text: str, err_text: str) -> Optional[str]:
    """從日誌內容推測較友善的錯誤訊息。"""
    joined = ((out_text or "") + "\n" + (err_text or "")).lower()
    if _log_contains_login_error(joined):
        return "學號或密碼錯誤，請重新輸入。"
    patterns = [
        ("nosuchframeexception", "無法切換到教務系統主畫面（main frame），網站結構可能變更，請稍後重試。"),
        ("找不到 main frame", "無法切換到教務系統主畫面（main frame），網站結構可能變更，請稍後重試。"),
        ("unable to locate element", "頁面元素找不到，可能網站改版或載入失敗，請重試或更新選擇器。"),
        ("no such element", "頁面元素找不到，可能網站改版或載入失敗，請重試或更新選擇器。"),
        ("timeoutexception", "操作逾時，請檢查網路或稍後再試。"),
        ("net::err", "網路連線錯誤，請檢查網路狀態或校務系統是否可連線。"),
        ("connection refused", "無法連線到網站，請稍後再試。"),
        ("this version of chromedriver only supports", "Chrome/Driver 版本不相容，請更新瀏覽器或驅動程式。"),
    ]
    for key, msg in patterns:
        if key in joined:
            return msg
    return None


def latest_existing(path_patterns):
    """
    回傳符合任一 pattern 的**最新**檔案路徑（找不到回傳 None）
    """
    candidates = []
    for pat in path_patterns:
        candidates.extend(glob.glob(pat))
    if not candidates:
        return None
    candidates.sort(key=lambda p: Path(p).stat().st_mtime, reverse=True)
    return candidates[0]


def load_csv_safely(path: str) -> pd.DataFrame:
    """
    嘗試用 UTF-8-SIG 讀，失敗就用 UTF-8。
    """
    try:
        return pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        return pd.read_csv(path, encoding="utf-8")


def filter_df(df: pd.DataFrame, keyword: str) -> pd.DataFrame:
    """
    針對所有欄位做簡單關鍵字包含過濾（不分大小寫）。
    """
    if not keyword:
        return df
    kw = str(keyword).strip().lower()
    mask = pd.Series([False]*len(df))
    for col in df.columns:
        mask = mask | df[col].astype(str).str.lower().str.contains(kw, na=False)
    return df[mask]


@app.route("/", methods=["GET"])
def index():
    return render_template("home.html")


@app.route("/query", methods=["POST"])
def query():
    kind = request.form.get("kind")  # timetable / grades / ranking / attendance
    keyword = request.form.get("keyword", "").strip()

    # 讀取現有環境變數
    current_user = os.getenv("SHU_USERNAME")
    current_pwd  = os.getenv("SHU_PASSWORD")

    # 表單傳入（可能為空）
    form_user = request.form.get("username") or None
    form_pwd  = request.form.get("password") or None

    # 以表單為優先，否則沿用現有環境
    user = form_user or current_user
    pwd  = form_pwd or current_pwd

    if not kind:
        flash("請選擇要查詢的類型")
        return redirect(url_for("index"))

    # 若表單提供的帳密與目前環境不同，則更新 .env 與目前行程的環境變數
    try:
        env_path = Path(".env")
        if user and (form_user is not None) and (form_user != current_user):
            set_key(str(env_path), "SHU_USERNAME", str(user))
            os.environ["SHU_USERNAME"] = str(user)
        if pwd is not None and (form_pwd is not None) and (form_pwd != current_pwd):
            set_key(str(env_path), "SHU_PASSWORD", str(pwd))
            os.environ["SHU_PASSWORD"] = str(pwd)
    except Exception as e:
        print(f"[ENV] 更新 .env 失敗：{e}")

    # 確定要使用的『使用者資料夾』
    # 1) 有填學號 -> 用該學號，並記錄到 .last_username
    # 2) 沒填學號 -> 若存在上一次學號則沿用；否則用環境中的學號
    effective_user = user
    if not form_user:
        if LAST_USER_FILE.exists():
            try:
                saved = LAST_USER_FILE.read_text(encoding="utf-8").strip()
                if saved:
                    effective_user = saved
            except Exception:
                pass
    else:
        try:
            LAST_USER_FILE.write_text(str(user), encoding="utf-8")
        except Exception:
            pass

    # 建立目錄：data/<username>
    work_dir = DATA_ROOT / (effective_user or "_unknown")
    work_dir.mkdir(parents=True, exist_ok=True)

    # 一律執行爬蟲以確保最新資料
    if not user or not pwd:
        flash("需要 SHU_USERNAME / SHU_PASSWORD 才能執行爬蟲")
        return redirect(url_for("index"))

    res = run_script(kind, {"SHU_USERNAME": user, "SHU_PASSWORD": pwd}, work_dir=work_dir)
    if res.get("code") != 0:
        if res.get("code") == 2:
            flash("學號或密碼錯誤，請重新輸入。", "danger")
            return redirect(url_for("index"))
        # 嘗試讀取輸出來判斷是否為登入錯誤
        try:
            out_text = Path(res.get("out", "")).read_text(encoding="utf-8", errors="ignore") if res.get("out") else ""
            err_text = Path(res.get("err", "")).read_text(encoding="utf-8", errors="ignore") if res.get("err") else ""
        except Exception:
            out_text = err_text = ""

        msg = _diagnose_message(out_text, err_text)
        if msg:
            flash(msg, "danger")
        else:
            # 顯示錯誤檔首行協助判讀
            first_err = (err_text or out_text or "").strip().splitlines()[:1]
            hint = f"（{first_err[0]}）" if first_err else ""
            flash(f"爬蟲執行失敗，請到 logs/ 夾查看 out/err 記錄 {hint}", "danger")
        return redirect(url_for("index"))

    # 決定要讀哪個 CSV
    outputs = OUTPUTS.get(kind, [])
    if not outputs:
        flash("沒有設定輸出的檔名，請檢查 app.py 的 OUTPUTS 設定")
        return redirect(url_for("index"))

    # 歷年成績有兩份 CSV：課程與彙總，優先顯示課程
    # 在使用者工作目錄底下找最新輸出
    user_patterns = [str((work_dir / Path(p)).as_posix()) for p in outputs]
    csv_path = latest_existing(user_patterns)
    if not csv_path:
        # 沒有產生 CSV，也檢查是否為登入錯誤
        try:
            out_text = Path(res.get("out", "")).read_text(encoding="utf-8", errors="ignore") if res.get("out") else ""
            err_text = Path(res.get("err", "")).read_text(encoding="utf-8", errors="ignore") if res.get("err") else ""
        except Exception:
            out_text = err_text = ""

        msg = _diagnose_message(out_text, err_text)
        if msg:
            flash(msg, "danger")
        else:
            # 顯示錯誤檔路徑，方便點開
            err_hint = res.get("err") or res.get("out")
            if err_hint:
                first_err = (err_text or out_text or "").strip().splitlines()[:1]
                hint = f"（{first_err[0]}）" if first_err else ""
                flash(f"找不到對應的輸出 CSV；請查看日誌：{err_hint} {hint}", "danger")
            else:
                flash("找不到對應的輸出 CSV，請先執行一次爬蟲或確認檔名", "danger")
        return redirect(url_for("index"))

    df = load_csv_safely(csv_path)
    df = filter_df(df, keyword)

    # 只挑常用欄位（有的話），避免表格太寬
    pref = DEFAULT_COLUMNS.get(kind, [])
    cols = [c for c in pref if c in df.columns]
    view_df = df[cols] if cols else df

    # 把目前顯示的 CSV 檔名也帶回前端（給下載）
    return render_template(
        "home.html",
        result_table=view_df.to_html(index=False, classes="table table-striped table-hover"),
        csv_path=csv_path,
        kind=kind,
        keyword=keyword
    )


@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not Path(path).exists():
        flash("檔案不存在")
        return redirect(url_for("index"))
    # 直接傳檔讓使用者下載
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    # python app.py
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=True)
