"""
MOPS PDF Packager — 批次下載模組

讀取「公司清單.xlsx」中的股票代碼與年份區間，
依序下載每家公司的所有 MOPS 報告（年報、財報、三書表、法說會簡報、ESG 永續報告書）。

功能：
  - 支援斷點續傳（透過 batch_progress.json 記錄已完成的公司）
  - 每家公司之間隨機等待 30~60 秒，避免被 MOPS 封鎖
  - 所有 log 同時輸出到 console 和 log 檔案
  - 跑完後印出 summary（成功/失敗統計）
"""

import sys
import os
import json
import time
import random
import datetime
import logging
import subprocess

# 將專案根目錄加入路徑
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.utils.downloader import MOPSDownloader


# ============================================================
# 設定
# ============================================================

EXCEL_FILENAME = "公司清單.xlsx"
PROGRESS_FILENAME = "batch_progress.json"
OUTPUT_DIR_NAME = "MOPS_批次下載"


# ============================================================
# Logging 設定
# ============================================================

def setup_logging(output_dir):
    """
    設定 logging：同時輸出到 console 和 log 檔案。
    Log 檔案存放在 output_dir 下。
    """
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f"batch_download_{timestamp}.log"
    log_path = os.path.join(output_dir, log_filename)

    # 建立 root logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # 清除既有 handler（避免重複）
    logger.handlers.clear()

    formatter = logging.Formatter(
        "[%(asctime)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # File handler
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # 將 print 的輸出也 capture 到 log 檔案
    # 透過覆寫 sys.stdout 讓現有 scraper 的 print() 也能寫入 log
    sys.stdout = _PrintLogger(logger, log_path)

    logging.info(f"Log 檔案: {log_path}")
    return logger


class _PrintLogger:
    """
    攔截 print() 輸出，同時寫入 console 和 log 檔案。
    讓現有 scraper 裡的 print() 也能被 capture。
    """
    def __init__(self, logger, log_path):
        self._logger = logger
        self._log_file = open(log_path, "a", encoding="utf-8")
        self._original_stdout = sys.__stdout__

    def write(self, message):
        if message and message.strip():
            self._original_stdout.write(message)
            if not message.endswith("\n"):
                self._original_stdout.write("\n")
            self._log_file.write(message)
            if not message.endswith("\n"):
                self._log_file.write("\n")
            self._log_file.flush()
        elif message == "\n":
            self._original_stdout.write(message)
            self._log_file.write(message)
            self._log_file.flush()

    def flush(self):
        self._original_stdout.flush()
        self._log_file.flush()

    def close(self):
        self._log_file.close()


# ============================================================
# Excel 讀取
# ============================================================

ALL_REPORT_TYPES = ["年報", "財報", "關係企業三書表", "法說會簡報", "ESG永續報告書"]

# Excel D~I 欄對應的報告類型（索引 3~8）
_REPORT_COL_MAP = {
    3: "全選",
    4: "年報",
    5: "財報",
    6: "關係企業三書表",
    7: "法說會簡報",
    8: "ESG永續報告書",
}


def read_company_list(excel_path):
    """
    讀取公司清單 Excel，回傳 list of dict。

    Excel 格式：
      A 欄: 股票代碼 (必填)
      B 欄: 起始年份 (民國年, 選填)
      C 欄: 結束年份 (民國年, 選填)
      D 欄: 全選 (✓ = 下載全部報告類型)
      E~I 欄: 年報 / 財報 / 關係企業三書表 / 法說會簡報 / ESG永續報告書 (✓ = 下載)

    Returns:
        list[dict]: [{"ticker": "2330", "year_start": 110, "year_end": 115,
                      "report_types": ["年報", "財報", ...]}, ...]
    """
    try:
        import openpyxl
    except ImportError:
        logging.error("缺少 openpyxl 套件。請執行: pip install openpyxl")
        sys.exit(1)

    if not os.path.exists(excel_path):
        logging.error(f"找不到公司清單: {excel_path}")
        logging.error("請先建立「公司清單.xlsx」，格式：A欄=股票代碼, B欄=起始年份(民國年), C欄=結束年份(民國年)")
        sys.exit(1)

    wb = openpyxl.load_workbook(excel_path, read_only=True)
    ws = wb.active

    companies = []
    current_roc_year = datetime.datetime.now().year - 1911

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or not row[0]:
            continue  # 跳過空列

        ticker = str(row[0]).strip()
        if not ticker:
            continue

        # 讀取起始/結束年份
        year_start = None
        year_end = None

        if len(row) > 1 and row[1] is not None:
            try:
                year_start = int(row[1])
            except (ValueError, TypeError):
                logging.warning(f"第 {row_idx} 列: 起始年份格式錯誤「{row[1]}」，將使用預設值")

        if len(row) > 2 and row[2] is not None:
            try:
                year_end = int(row[2])
            except (ValueError, TypeError):
                logging.warning(f"第 {row_idx} 列: 結束年份格式錯誤「{row[2]}」，將使用預設值")

        # 預設值邏輯
        if year_start is None and year_end is None:
            year_end = current_roc_year
            year_start = current_roc_year - 4
        elif year_start is not None and year_end is None:
            year_end = current_roc_year
        elif year_start is None and year_end is not None:
            year_start = year_end - 4

        # 讀取報告類型 (D~I 欄, 索引 3~8)
        def _is_checked(col_idx):
            if len(row) > col_idx and row[col_idx]:
                val = str(row[col_idx]).strip()
                return val in ("\u2713", "V", "v", "O", "o", "Y", "y", "1", "TRUE", "True", "true")
            return False

        select_all = _is_checked(3)  # D 欄: 全選

        if select_all:
            report_types = list(ALL_REPORT_TYPES)
        else:
            report_types = []
            for col_idx in range(4, 9):  # E~I 欄
                if _is_checked(col_idx):
                    report_types.append(_REPORT_COL_MAP[col_idx])

            # 如果 D~I 全部留空，當作全選（防呆）
            if not report_types:
                report_types = list(ALL_REPORT_TYPES)

        companies.append({
            "ticker": ticker,
            "year_start": year_start,
            "year_end": year_end,
            "report_types": report_types,
        })

    wb.close()
    return companies


# ============================================================
# 斷點續傳
# ============================================================

def load_progress(progress_path):
    """載入已完成的 ticker 清單"""
    if os.path.exists(progress_path):
        try:
            with open(progress_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return set(data.get("completed", []))
        except (json.JSONDecodeError, KeyError):
            return set()
    return set()


def save_progress(progress_path, completed_set):
    """儲存已完成的 ticker 清單"""
    with open(progress_path, "w", encoding="utf-8") as f:
        json.dump({"completed": sorted(completed_set)}, f, ensure_ascii=False, indent=2)


# ============================================================
# 自動 Push 到 GitHub
# ============================================================

def auto_push_to_github(script_dir):
    """
    下載完成後自動將報告 push 到 GitHub。
    如果 git 未安裝或不是 git repo，會顯示提示並跳過。
    """
    logging.info("\n" + "=" * 60)
    logging.info("自動上傳到 GitHub")
    logging.info("=" * 60)

    # 1. 檢查 git 是否已安裝
    try:
        subprocess.run(
            ["git", "--version"],
            cwd=script_dir, capture_output=True, check=True
        )
    except FileNotFoundError:
        logging.warning("未偵測到 git，跳過自動上傳。")
        logging.warning("如需自動上傳功能，請先安裝 Git: https://git-scm.com/")
        return
    except subprocess.CalledProcessError:
        logging.warning("git 執行異常，跳過自動上傳。")
        return

    # 2. 檢查是否為 git repo
    result = subprocess.run(
        ["git", "rev-parse", "--git-dir"],
        cwd=script_dir, capture_output=True
    )
    if result.returncode != 0:
        logging.warning("此目錄不是 Git 儲存庫，跳過自動上傳。")
        logging.warning("請先執行 git init 並設定 remote。")
        return

    # 3. git add
    logging.info("正在加入檔案到 Git...")
    subprocess.run(
        ["git", "add", "MOPS_批次下載/"],
        cwd=script_dir, capture_output=True
    )
    subprocess.run(
        ["git", "add", "公司清單.xlsx"],
        cwd=script_dir, capture_output=True
    )

    # 4. 檢查是否有變更需要 commit
    status_result = subprocess.run(
        ["git", "diff", "--cached", "--quiet"],
        cwd=script_dir, capture_output=True
    )
    if status_result.returncode == 0:
        logging.info("沒有新的變更需要上傳。")
        return

    # 5. git commit
    today = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    commit_msg = f"更新報告 {today}"
    logging.info(f"正在 commit: {commit_msg}")
    commit_result = subprocess.run(
        ["git", "commit", "-m", commit_msg],
        cwd=script_dir, capture_output=True, text=True
    )
    if commit_result.returncode != 0:
        logging.error(f"git commit 失敗: {commit_result.stderr}")
        return

    # 6. git push
    logging.info("正在 push 到 GitHub...")
    push_result = subprocess.run(
        ["git", "push"],
        cwd=script_dir, capture_output=True, text=True
    )
    if push_result.returncode == 0:
        logging.info("成功上傳到 GitHub！")
    else:
        logging.error(f"git push 失敗: {push_result.stderr}")
        logging.error("請確認網路連線及 GitHub 認證設定。")


# ============================================================
# 主流程
# ============================================================

def main():
    # PyInstaller --onefile 會解壓到暫存目錄，__file__ 指向暫存路徑
    # 必須用 sys.executable 取得 exe 實際所在資料夾
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(script_dir, EXCEL_FILENAME)
    progress_path = os.path.join(script_dir, PROGRESS_FILENAME)
    output_dir = os.path.join(script_dir, OUTPUT_DIR_NAME)

    # 設定 logging
    logger = setup_logging(output_dir)

    logging.info("=" * 60)
    logging.info("MOPS PDF Packager — 批次下載")
    logging.info("=" * 60)

    # 讀取公司清單
    companies = read_company_list(excel_path)
    if not companies:
        logging.error("公司清單為空，請確認 Excel 內容。")
        sys.exit(1)

    logging.info(f"共讀取到 {len(companies)} 家公司")
    for c in companies:
        types_str = ", ".join(c["report_types"])
        logging.info(f"  {c['ticker']}  民國 {c['year_start']}~{c['year_end']} 年  [{types_str}]")

    # 載入進度
    completed = load_progress(progress_path)
    if completed:
        logging.info(f"偵測到先前進度，已完成 {len(completed)} 家: {', '.join(sorted(completed))}")

    # 開始下載
    success_list = []
    fail_list = []

    for idx, company in enumerate(companies, start=1):
        ticker = company["ticker"]

        if ticker in completed:
            logging.info(f"\n[{idx}/{len(companies)}] {ticker} — 已完成，跳過")
            success_list.append(ticker)
            continue

        logging.info(f"\n{'=' * 60}")
        logging.info(f"[{idx}/{len(companies)}] 開始下載 {ticker} (民國 {company['year_start']}~{company['year_end']} 年)")
        logging.info("=" * 60)

        try:
            downloader = MOPSDownloader(
                ticker=ticker,
                save_base_dir=output_dir,
                year_start=company["year_start"],
                year_end=company["year_end"],
                report_types=company["report_types"],
            )
            downloader.run(use_subdirs=True)

            # 公開說明書：MOPS 目前未提供公開說明書的下載介面
            logging.info(f"[{ticker}] 注意: MOPS 目前未提供公開說明書的自動下載介面，此類型已跳過。")

            # 記錄成功
            success_list.append(ticker)
            completed.add(ticker)
            save_progress(progress_path, completed)
            logging.info(f"[{ticker}] 下載完成！已記錄進度。")

        except Exception as e:
            logging.error(f"[{ticker}] 下載過程發生錯誤: {e}")
            fail_list.append(ticker)

        # 如果不是最後一家，等待一段時間
        if idx < len(companies):
            wait_seconds = random.uniform(45, 90)
            logging.info(f"\n等待 {wait_seconds:.0f} 秒後繼續下一家...")
            time.sleep(wait_seconds)

    # ============================================================
    # Summary
    # ============================================================
    logging.info("\n" + "=" * 60)
    logging.info("批次下載完成 — Summary")
    logging.info("=" * 60)
    logging.info(f"成功: {len(success_list)} 家")
    logging.info(f"失敗: {len(fail_list)} 家")

    if fail_list:
        logging.info(f"失敗的公司: {', '.join(fail_list)}")

    logging.info(f"\n檔案存放於: {os.path.abspath(output_dir)}")
    logging.info(f"進度檔案: {progress_path}")
    logging.info("如需重新下載失敗的公司，請從 batch_progress.json 中移除對應代碼後重跑。")

    # 自動上傳到 GitHub
    auto_push_to_github(script_dir)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n發生未預期的錯誤: {e}")
    finally:
        print()
        input("按 Enter 鍵關閉視窗...")
