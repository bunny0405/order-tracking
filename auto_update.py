"""
追單到貨追蹤系統 - 自動更新排程
設定在 Cowork 排程：每週三、每週五執行

執行前請確認：
  1. C:\Users\bunny\Desktop\追加到貨追蹤\ 內有最新的 Excel 檔案
  2. 電腦已安裝 Git 並設定好 GitHub 帳號
"""

import os
import glob
import subprocess
from datetime import datetime

BASE_DIR = r"C:\Users\bunny\Desktop\追加到貨追蹤"
GENERATE = os.path.join(BASE_DIR, "generate.py")
LOG_FILE = os.path.join(BASE_DIR, "update_log.txt")

def log(msg):
    ts = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(line + '\n')

def run(cmd, cwd=BASE_DIR):
    result = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, encoding='utf-8')
    if result.returncode != 0:
        raise RuntimeError(f"指令失敗: {' '.join(cmd)}\n{result.stderr}")
    return result.stdout.strip()

def check_files():
    """確認必要的 Excel 檔案都存在"""
    required = ['追加單_*.xlsx', '採購總表_*.xlsx', '調撥單_*.xlsx', '門市基本資料*.xlsx']
    missing = []
    for pattern in required:
        if not glob.glob(os.path.join(BASE_DIR, pattern)):
            missing.append(pattern)
    if missing:
        raise FileNotFoundError(f"找不到以下檔案：{', '.join(missing)}\n請先把最新 Excel 放到資料夾：{BASE_DIR}")

# ── 主流程 ────────────────────────────────────────────────────────────────
log("=" * 50)
log("開始執行自動更新")

try:
    # 1. 確認檔案
    log("檢查 Excel 檔案...")
    check_files()
    log("✅ 檔案確認完成")

    # 2. 執行 generate.py 產出 index.html
    log("執行 generate.py...")
    output = run(['python', GENERATE], cwd=BASE_DIR)
    log(f"✅ 產出完成\n{output}")

    # 3. Git add & commit
    today = datetime.now().strftime('%Y%m%d_%H%M')
    log("推送至 GitHub...")
    run(['git', 'add', 'index.html'])
    run(['git', 'commit', '-m', f'update {today}'])
    run(['git', 'push'])
    log("✅ GitHub Pages 已更新")

    log(f"🎉 全部完成！網址：https://bunny0405.github.io/order-tracking/")

except FileNotFoundError as e:
    log(f"❌ 檔案錯誤：{e}")
    raise
except Exception as e:
    log(f"❌ 執行失敗：{e}")
    raise

log("=" * 50)
