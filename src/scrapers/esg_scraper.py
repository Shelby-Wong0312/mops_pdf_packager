"""
ESG 永續報告書 下載模組
從 ESG 數位平台 (esggenplus.twse.com.tw) 搜尋並下載永續報告書 PDF。

API 端點:
  POST /api/api/MopsSustainReport/data
  Body: {"companyCodeList":["2330"],"year":2024,"industryNameList":[],
         "marketType":0,"industryName":"all","companyCode":"2330"}
  marketType: 0=上市, 1=上櫃

  GET  /api/api/MopsSustainReport/data/FileStream?id={GUID}
"""

import requests
import os
import time
import random
import datetime
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://esggenplus.twse.com.tw"
API_BASE = f"{BASE_URL}/api/api"


def _create_session():
    """建立帶有正確 headers 的 requests session，並先訪問主頁取得 cookies"""
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/131.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
        "Content-Type": "application/json",
        "Referer": f"{BASE_URL}/inquiry/report",
        "Origin": BASE_URL,
    })

    # 先訪問主頁，讓 server 設定 session cookies (例如 .AspNetCore.Antiforgery)
    try:
        print("[ESG] 正在連線 ESG 數位平台...")
        home_res = session.get(BASE_URL, verify=False, timeout=15, allow_redirects=True)
        print(f"[ESG] 主頁回應: HTTP {home_res.status_code}")
    except Exception as e:
        print(f"[ESG] Warning: 無法訪問 ESG 平台主頁: {e}")

    return session


def _get_antiforgery_token(session):
    """
    取得 CSRF / Antiforgery token。
    API 回傳格式: {"count":1, "data":"CfDJ8...", "code":200, "success":true}
    Token 在 "data" 欄位中。
    """
    try:
        res = session.get(f"{API_BASE}/Antiforgery/token", verify=False, timeout=10)
        res.raise_for_status()
        result = res.json()

        # Token 藏在 "data" 欄位
        token = result.get("data", "")
        if isinstance(token, str) and len(token) > 20:
            session.headers["RequestVerificationToken"] = token
            print(f"[ESG] 取得 Antiforgery token OK (長度: {len(token)})")
            return token
        else:
            print(f"[ESG] Warning: Antiforgery token 格式異常: {str(result)[:100]}")
            return ""
    except Exception as e:
        print(f"[ESG] Warning: 取得 Antiforgery token 失敗: {e}")
        return ""


def _search_esg_reports(session, ticker, target_year=None, max_reports=3):
    """
    搜尋指定公司的永續報告書列表。
    如果指定 target_year，只搜尋該年份。
    否則搜尋最近 max_reports 年。
    """
    ticker_str = str(ticker)
    url = f"{API_BASE}/MopsSustainReport/data"

    # 決定要搜哪幾年
    if target_year:
        search_years = [int(target_year)]
    else:
        current_year = datetime.datetime.now().year
        # 只搜最近幾年，不要搜到 2010
        search_years = list(range(current_year, current_year - max_reports - 1, -1))
        print(f"[ESG] 搜尋年份: {search_years}")

    all_reports = []

    for search_year in search_years:
        found = False
        # 嘗試 marketType=0 (上市) 和 marketType=1 (上櫃)
        for market_type in [0, 1]:
            try:
                payload = {
                    "companyCodeList": [ticker_str],
                    "year": search_year,
                    "industryNameList": [],
                    "marketType": market_type,
                    "industryName": "all",
                    "companyCode": ticker_str,
                }

                print(f"[ESG] 查詢 {ticker_str} / {search_year} / marketType={market_type}...")
                res = session.post(url, json=payload, verify=False, timeout=20)
                print(f"[ESG]   -> HTTP {res.status_code}")

                if res.status_code == 200:
                    data = res.json()
                    success = data.get("success", False)
                    msg = data.get("message", "")
                    items = data.get("data", [])

                    if success and items:
                        for item in items:
                            tw_id = item.get("twFirstReportDownloadId", "")
                            if tw_id == "00000000-0000-0000-0000-000000000000":
                                tw_id = ""
                            en_id = item.get("enFirstReportDownloadId", "")
                            if en_id == "00000000-0000-0000-0000-000000000000":
                                en_id = ""

                            report = {
                                "year": search_year,
                                "code": item.get("code", ticker_str),
                                "name": item.get("shortName", ""),
                                "tw_doc_link": item.get("twDocLink", ""),
                                "tw_download_id": tw_id,
                                "en_doc_link": item.get("enDocLink", ""),
                                "en_download_id": en_id,
                                "reporting_interval": item.get("reportingInterval", ""),
                                "market_type": market_type,
                            }
                            all_reports.append(report)
                            print(f"[ESG]   找到: {report['name']} ({report['code']}) - {search_year}")
                            if report["tw_doc_link"]:
                                print(f"[ESG]   PDF直連: {report['tw_doc_link'][:80]}...")
                            if report["tw_download_id"]:
                                print(f"[ESG]   下載GUID: {report['tw_download_id'][:16]}...")

                        found = True
                        break  # 找到了，不用再試另一個 marketType
                    else:
                        if "查無" in msg:
                            pass  # 正常，不是這個 marketType
                        else:
                            print(f"[ESG]   API回應: success={success}, msg={msg}")

                elif res.status_code == 400:
                    print(f"[ESG]   400 Error: {res.text[:200]}")

            except requests.exceptions.ConnectionError as e:
                print(f"[ESG]   連線失敗: {e}")
                return all_reports  # 網路有問題，不要繼續了
            except Exception as e:
                print(f"[ESG]   搜尋失敗: {e}")

        if not found:
            print(f"[ESG]   {search_year} 年查無 {ticker_str} 永續報告")

        time.sleep(random.uniform(0.3, 1.0))

    return all_reports


def _download_pdf(url, save_path, label="", timeout=120):
    """通用 PDF 下載函式，回傳是否成功"""
    try:
        dl_session = requests.Session()
        dl_session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0.0.0 Safari/537.36"
            ),
        })
        res = dl_session.get(url, verify=False, timeout=timeout, stream=True)
        print(f"[ESG]   {label} HTTP {res.status_code}, Content-Type: {res.headers.get('Content-Type', 'N/A')[:50]}")

        if res.status_code != 200:
            print(f"[ESG]   {label} 非 200 回應，跳過。")
            return False

        # 寫入檔案
        with open(save_path, "wb") as f:
            for chunk in res.iter_content(chunk_size=8192):
                f.write(chunk)

        file_size = os.path.getsize(save_path)
        if file_size > 10000:  # 至少 10KB
            file_size_mb = file_size / (1024 * 1024)
            print(f"[ESG]   下載成功! {file_size_mb:.1f} MB [{label}]")
            return True
        else:
            # 檔案太小，可能是錯誤頁面
            # 印出前 200 bytes 方便 debug
            with open(save_path, "rb") as f:
                head = f.read(200)
            print(f"[ESG]   檔案太小 ({file_size} bytes): {head[:100]}")
            os.remove(save_path)
            return False

    except Exception as e:
        print(f"[ESG]   {label} 下載失敗: {e}")
        if os.path.exists(save_path):
            os.remove(save_path)
        return False


def _download_esg_pdf(session, report, save_dir):
    """
    下載永續報告書 PDF，嘗試三種方式:
    1. twDocLink (公司自家網站 PDF 直連)
    2. FileStream API (ESG 平台 GUID)
    3. 直接用 session 帶 cookies 的 FileStream
    """
    ticker = report["code"]
    year = report["year"]
    tw_doc_link = report.get("tw_doc_link", "")
    tw_download_id = report.get("tw_download_id", "")

    os.makedirs(save_dir, exist_ok=True)
    filename = f"{ticker}_{year}_永續報告書.pdf"
    save_path = os.path.join(save_dir, filename)

    # 檔案已存在且夠大，跳過
    if os.path.exists(save_path) and os.path.getsize(save_path) > 10000:
        print(f"[ESG] {filename} 已存在，跳過。")
        return save_path

    print(f"[ESG] 下載 {filename}...")

    # 方法 1: twDocLink 公司自家網站
    if tw_doc_link:
        print(f"[ESG]   方法1: 公司網站直連")
        if _download_pdf(tw_doc_link, save_path, label="公司網站"):
            return save_path

    # 方法 2: FileStream (新 session，不帶 ESG cookies)
    if tw_download_id:
        download_url = f"{API_BASE}/MopsSustainReport/data/FileStream?id={tw_download_id}"
        print(f"[ESG]   方法2: ESG平台 FileStream (新session)")
        if _download_pdf(download_url, save_path, label="FileStream-新session"):
            return save_path

    # 方法 3: FileStream (用原本帶 cookies 的 session)
    if tw_download_id:
        download_url = f"{API_BASE}/MopsSustainReport/data/FileStream?id={tw_download_id}"
        print(f"[ESG]   方法3: ESG平台 FileStream (帶cookies)")
        try:
            res = session.get(download_url, verify=False, timeout=120, stream=True)
            print(f"[ESG]   帶cookies HTTP {res.status_code}")
            if res.status_code == 200:
                with open(save_path, "wb") as f:
                    for chunk in res.iter_content(chunk_size=8192):
                        f.write(chunk)
                file_size = os.path.getsize(save_path)
                if file_size > 10000:
                    print(f"[ESG]   下載成功! {file_size/(1024*1024):.1f} MB [帶cookies]")
                    return save_path
                else:
                    os.remove(save_path)
        except Exception as e:
            print(f"[ESG]   帶cookies下載失敗: {e}")
            if os.path.exists(save_path):
                os.remove(save_path)

    print(f"[ESG] *** 無法下載 {filename} (所有方式均失敗) ***")
    return None


def download_esg_report(ticker, save_dir, max_reports=3):
    """
    從 ESG 數位平台下載永續報告書 PDF。

    Args:
        ticker: 股票代碼 (例如 "2330")
        save_dir: 儲存目錄
        max_reports: 最多下載幾份報告 (預設 3，同時也是搜尋最近幾年)

    Returns:
        list: 成功下載的檔案路徑列表
    """
    ticker_str = str(ticker)
    print(f"\n{'='*60}")
    print(f"[ESG] 開始搜尋 {ticker_str} 永續報告書 (ESG 數位平台)")
    print(f"[ESG] 平台: {BASE_URL}")
    print(f"{'='*60}")

    # Step 1: 建立 session (含 cookies)
    session = _create_session()
    time.sleep(random.uniform(0.5, 1.0))

    # Step 2: 取得 antiforgery token
    _get_antiforgery_token(session)
    time.sleep(random.uniform(0.3, 0.8))

    # Step 3: 搜尋永續報告
    reports = _search_esg_reports(session, ticker_str, max_reports=max_reports)

    if not reports:
        print(f"\n[ESG] 查無 {ticker_str} 的永續報告書。")
        print(f"[ESG] 可手動至以下網址搜尋:")
        print(f"[ESG]   {BASE_URL}/inquiry/report?companyCode={ticker_str}")
        return []

    # 按年份由新到舊排序
    reports.sort(key=lambda x: x["year"], reverse=True)
    print(f"\n[ESG] 共找到 {len(reports)} 份報告，準備下載最新 {min(max_reports, len(reports))} 份...")

    # Step 4: 下載 PDF
    saved_paths = []
    for report in reports[:max_reports]:
        time.sleep(random.uniform(1, 2))
        path = _download_esg_pdf(session, report, save_dir)
        if path:
            saved_paths.append(path)

    print(f"\n[ESG] === 完成! 成功下載 {len(saved_paths)}/{min(max_reports, len(reports))} 份永續報告書 ===")
    return saved_paths


if __name__ == "__main__":
    import sys
    ticker = sys.argv[1] if len(sys.argv) > 1 else "2330"
    output_dir = sys.argv[2] if len(sys.argv) > 2 else f"./{ticker}_ESG_Reports"
    paths = download_esg_report(ticker, output_dir)
    if paths:
        print(f"\n下載完成! 檔案位置:")
        for p in paths:
            print(f"  {p}")
    else:
        print(f"\n未能下載任何報告。")
