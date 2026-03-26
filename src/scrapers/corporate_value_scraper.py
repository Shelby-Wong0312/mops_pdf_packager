import requests
from bs4 import BeautifulSoup
import re
import os
import urllib3
import time
import random

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

MAX_RETRIES = 3
RETRY_WAIT = 60  # 秒


def download_corporate_value_pdf(ticker, year, save_dir, download_all=False):
    """
    從 MOPS 下載提升企業價值計劃 PDF。
    資料來源: https://mopsov.twse.com.tw/mops/web/t100sb16

    Args:
        ticker: 股票代碼
        year: 民國年
        save_dir: 儲存目錄
        download_all: True 回傳所有下載路徑的 list，False 回傳第一筆路徑或 None

    Returns:
        list[str] if download_all=True, else str or None
    """
    url = "https://mopsov.twse.com.tw/mops/web/ajax_t100sb16"

    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    }

    saved_paths = []

    # 嘗試上市 (sii) 和上櫃 (otc)
    for mar_kind in ["sii", "otc"]:
        payload = {
            "encodeURIComponent": "1",
            "step": "1",
            "firstin": "1",
            "TYPEK": mar_kind,
            "MAR_KIND": mar_kind,
            "STATUS": "Y",
            "EYEAR": str(year),
            "CO_ID": str(ticker),
            "BCODE": "",
        }

        try:
            time.sleep(random.uniform(5, 10))

            res = None
            for attempt in range(MAX_RETRIES):
                res = session.post(url, data=payload, headers=headers, verify=False, timeout=15)
                res.raise_for_status()

                blocked_keywords = ["查詢過量", "SECURITY", "ACCESSED", "請稍後再查詢"]
                if any(kw in res.text for kw in blocked_keywords):
                    print(f"[{ticker}] 提升企業價值計劃查詢遭伺服器封鎖，等待 {RETRY_WAIT} 秒後重試 ({attempt + 1}/{MAX_RETRIES})...")
                    time.sleep(RETRY_WAIT)
                    continue
                break
            else:
                print(f"[{ticker}] 提升企業價值計劃查詢持續被封鎖，跳過 {year} 年度 ({mar_kind})。")
                continue

            soup = BeautifulSoup(res.content, "html.parser")

            # 找 FileDownLoad 的 onclick 按鈕
            # 格式: window.open("/server-java/FileDownLoad?step=9&fileName=XXXX.pdf&filePath=/home/html/nas/protect/t100/", ...)
            file_buttons = soup.find_all(
                "input", type="button",
                onclick=re.compile(r"FileDownLoad.*fileName.*\.pdf", re.IGNORECASE)
            )

            if not file_buttons:
                # 也嘗試找 <a> 或其他元素
                file_buttons = soup.find_all(
                    attrs={"onclick": re.compile(r"FileDownLoad.*fileName.*\.pdf", re.IGNORECASE)}
                )

            if not file_buttons:
                continue  # 此市場類型沒找到，試下一個

            target_buttons = file_buttons if download_all else file_buttons[:1]

            for btn in target_buttons:
                onclick = btn.get("onclick", "")

                # 提取 fileName 和 filePath
                fn_match = re.search(r"fileName=([^&\"']+)", onclick)
                fp_match = re.search(r"filePath=([^&\"']+)", onclick)

                if not fn_match:
                    continue

                file_name = fn_match.group(1)
                file_path = fp_match.group(1) if fp_match else "/home/html/nas/protect/t100/"

                time.sleep(random.uniform(4, 7))
                print(f"找到 {ticker} 提升企業價值計劃 {file_name}，準備下載...")

                download_url = "https://mopsov.twse.com.tw/server-java/FileDownLoad"
                dl_params = {
                    "step": "9",
                    "fileName": file_name,
                    "filePath": file_path,
                }

                dl_res = session.get(download_url, params=dl_params, headers=headers, verify=False, timeout=30)
                dl_res.raise_for_status()

                content_type = dl_res.headers.get("Content-Type", "").lower()
                if "pdf" in content_type or len(dl_res.content) > 10000:
                    os.makedirs(save_dir, exist_ok=True)

                    # 從 fileName 中提取日期
                    date_match = re.search(r"(\d{8})", file_name)
                    date_str = date_match.group(1) if date_match else str(year)

                    readable_filename = f"{ticker}_{date_str}_提升企業價值計劃.pdf"
                    save_path = os.path.join(save_dir, readable_filename)
                    with open(save_path, "wb") as f:
                        f.write(dl_res.content)
                    print(f"已成功儲存 {readable_filename}\n")
                    saved_paths.append(save_path)

            if saved_paths:
                break  # 已找到，不用再試另一個市場

        except Exception as e:
            print(f"Error downloading corporate value plan for {ticker} ({mar_kind}): {e}")
            continue

    if not saved_paths:
        print(f"Warning: 找不到 {ticker} 在 {year} 年度的提升企業價值計劃。")

    return saved_paths if download_all else (saved_paths[0] if saved_paths else None)


if __name__ == "__main__":
    download_corporate_value_pdf("2330", 113, "./test_pack")
