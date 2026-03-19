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


def download_briefing_pdf(ticker, year, save_dir, download_all=False):
    """
    從 MOPS 下載法說會簡報 PDF。
    包含 retry 機制：遇到「查詢過量」或安全性封鎖時等待 60 秒後重試，最多 3 次。
    """
    url = "https://mopsov.twse.com.tw/mops/web/ajax_t100sb02_1"

    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    }

    # 查詢法說會，給定公司代號與年度
    payload = {
        "encodeURIComponent": "1",
        "step": "1",
        "firstin": "1",
        "off": "1",
        "queryName": "co_id",
        "inpuType": "co_id",
        "TYPEK": "all",
        "isnew": "false", # false 表示查歷史
        "co_id": str(ticker),
        "year": str(year)
    }

    try:
        time.sleep(random.uniform(5, 10)) # 搜尋前加入延遲，防止 mops.twse 阻擋

        res = None
        for attempt in range(MAX_RETRIES):
            res = session.post(url, data=payload, headers=headers, verify=False, timeout=15)
            res.raise_for_status()

            blocked_keywords = ["查詢過量", "SECURITY", "ACCESSED", "請稍後再查詢"]
            if any(kw in res.text for kw in blocked_keywords):
                print(f"[{ticker}] 法說會查詢遭伺服器封鎖，等待 {RETRY_WAIT} 秒後重試 ({attempt + 1}/{MAX_RETRIES})...")
                time.sleep(RETRY_WAIT)
                continue
            break  # 沒被封鎖
        else:
            print(f"[{ticker}] 法說會查詢持續被封鎖，跳過 {year} 年度。")
            return [] if download_all else None
            
        # 尋找第一種結構: 新年度法說會常出現的 document.fm...
        soup = BeautifulSoup(res.content, "html.parser")
        buttons = soup.find_all('input', type='button', onclick=re.compile(r"document\.fm\S*\.step\.value='9'"))
        
        # 尋找第二種結構: 歷史法說會常見的 form_fileDownload (a tag)
        # 例如: document.fm_fileDownload.fileName.value="245420240426M001.pdf"
        file_matches = re.findall(r'document\.fm_fileDownload\.fileName\.value=["\']([^"\']+)["\']', res.text)
        
        if not buttons and not file_matches:
            print(f"Warning: 找不到 {ticker} 在 {year} 年度的法說會簡報。 回傳的內容開頭為: {res.text[:300]}")
            return [] if download_all else None
            
        saved_paths = []
        
        # 處理第一種結構
        target_buttons = buttons if download_all else buttons[:1]
        for btn in target_buttons:
            time.sleep(random.uniform(4, 7)) # 下載前延遲
            onclick_script = btn['onclick']
            params = {}
            for match in re.finditer(r"document\.(\w+)\.(\w+)\.value='([^']*)'", onclick_script):
                field = match.group(2)
                value = match.group(3)
                params[field] = value
                
            if not params:
                continue
                
            print(f"找到 {ticker} 法說會簡報參數 {params.get('S_DAT', '')}，準備下載...")
            params['step'] = '9'
            
            download_res = session.post(url, data=params, headers=headers, verify=False, timeout=30)
            download_res.raise_for_status()
            
            content_type = download_res.headers.get("Content-Type", "").lower()
            if "pdf" in content_type or "powerpoint" in content_type or "presentation" in content_type or len(download_res.content) > 10000:
                os.makedirs(save_dir, exist_ok=True)
                
                # 判斷副檔名
                ext = ".pdf"
                cd = download_res.headers.get("Content-Disposition", "")
                m = re.search(r'filename="?([^";\s]+)"?', cd)
                if m:
                    guessed_ext = os.path.splitext(m.group(1))[1].lower()
                    if guessed_ext:
                        ext = guessed_ext
                elif "powerpoint" in content_type or "presentation" in content_type:
                    ext = ".pptx" if "presentationml" in content_type else ".ppt"
                    
                readable_filename = f"{ticker}_法說會簡報_{params.get('S_DAT', 'latest')}{ext}"
                save_path = os.path.join(save_dir, readable_filename)
                with open(save_path, "wb") as f:
                    f.write(download_res.content)
                print(f"已成功儲存 {readable_filename}\n")
                saved_paths.append(save_path)
                
        # 處理第二種結構 (且優先抓取包含 'M' 中文的檔案，過濾英文 'E')
        # 去除重複
        unique_files = list(dict.fromkeys(file_matches))
        target_files = []
        for f in unique_files:
            if "M001" in f or "M002" in f or "M" in f:
                target_files.append(f)
        if not target_files:
            target_files = unique_files # 如果沒有中文就全抓
            
        target_files = target_files if download_all else target_files[:1]
        
        for file_name in target_files:
            time.sleep(random.uniform(4, 7))
            print(f"找到 {ticker} 歷史法說會簡報 {file_name}，準備透過 FileDownLoad 下載...")
            download_url = "https://mopsov.twse.com.tw/server-java/FileDownLoad"
            dl_payload = {
                'step': '9',
                'filePath': '/home/html/nas/STR/',
                'functionName': 't100sb02_1',
                'fileName': file_name
            }
            dl_res = session.post(download_url, data=dl_payload, headers=headers, verify=False, timeout=30)
            dl_res.raise_for_status()
            
            # 檔案名稱中若是包含了年月日，我們可以取前 8 碼作為時間特徵，如 245420231027M001 -> 20231027
            time_str = "歷史"
            m = re.search(r'\d{4}(\d{8})', file_name)
            if m:
                time_str = m.group(1)
            
            if "pdf" in dl_res.headers.get("Content-Type", "").lower() or len(dl_res.content) > 10000:
                os.makedirs(save_dir, exist_ok=True)
                
                ext = os.path.splitext(file_name)[1].lower()
                if not ext:
                    ext = ".pdf"
                    
                readable_filename = f"{ticker}_法說會簡報_{time_str}{ext}"
                save_path = os.path.join(save_dir, readable_filename)
                with open(save_path, "wb") as f:
                    f.write(dl_res.content)
                print(f"已成功儲存 {readable_filename}\n")
                saved_paths.append(save_path)
                
        return saved_paths if download_all else (saved_paths[0] if saved_paths else None)
        
        
    except Exception as e:
        print(f"Error downloading briefing for {ticker}: {e}")
        return [] if download_all else None

if __name__ == "__main__":
    download_briefing_pdf("2330", "./test_pack")
