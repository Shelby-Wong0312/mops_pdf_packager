import os
import time
import re
from bs4 import BeautifulSoup
import glob

def init_driver(save_dir):
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.options import Options
    
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    
    # 設定下載資料夾
    os.makedirs(save_dir, exist_ok=True)
    abs_save_dir = os.path.abspath(save_dir)
    prefs = {
        "download.default_directory": abs_save_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True # 防止 Chrome 直接打開 PDF
    }
    options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver, abs_save_dir

def wait_for_new_file(abs_save_dir, old_files, timeout=30):
    for _ in range(timeout):
        current_files = set(os.listdir(abs_save_dir))
        new_files = current_files - old_files
        
        # 確認不是 .crdownload 暫存檔
        if new_files:
            for f in new_files:
                if not f.endswith(".crdownload") and not f.endswith(".tmp"):
                    return f
        time.sleep(1)
    return None

def download_briefing_selenium(ticker, save_dir):
    """
    從 mopsov 下載最近一次的法說會簡報 PDF
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    print(f"啟動瀏覽器準備下載 {ticker} 法說會簡報...")
    driver, abs_save_dir = init_driver(save_dir)
    
    try:
        url = "https://mopsov.twse.com.tw/mops/web/t146sb08"
        driver.get(url)
        
        # 等待輸入框出現 (尋找真正表單內的 co_id，避開全站搜尋的 keyword)
        co_id_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='text' and @name='co_id']"))
        )
        co_id_input.clear()
        co_id_input.send_keys(str(ticker))
        
        # 點擊真正的查詢按鈕 (避開全站搜尋)
        driver.execute_script("""
            var btns = document.querySelectorAll('input[type="button"]');
            var clicked = false;
            for(var i=0; i<btns.length; i++) {
                var oc = btns[i].getAttribute('onclick') || '';
                var nm = btns[i].getAttribute('name') || '';
                if(oc.indexOf('doAction') !== -1 || oc.indexOf('ajax') !== -1 || nm === 'rulesubmit') {
                    btns[i].click();
                    clicked = true;
                    break;
                }
            }
            if (!clicked && btns.length > 0) {
                btns[btns.length-1].click();
            }
        """)
        
        print(f"已送出法說會查詢，等待表格渲染...")
        time.sleep(3)
        
        # 利用 XPath 找出包含「簡報」二字的那一列 `<tr>`，並在該列中找到「檔案下載」按鈕
        # //table[@class='hasBorder']//tr[contains(., '簡報')]//input[@value='檔案下載']
        
        try:
            # 等待表格出現
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//table[@class='hasBorder']"))
            )
            
            # 使用 JS 找出含有「簡報」列的下載按鈕並點擊
            found = driver.execute_script("""
                var rows = document.querySelectorAll('table.hasBorder tr');
                for (var i = 0; i < rows.length; i++) {
                    // 如果該列內文有 '簡報' (或類似)，且裡面包含按鈕
                    if (rows[i].innerText.indexOf('簡報') !== -1 || rows[i].innerText.indexOf('r') !== -1) {
                        var btns = rows[i].querySelectorAll('input[type="button"]');
                        for (var j = 0; j < btns.length; j++) {
                            // 檔案下載按鈕
                            if (btns[j].value.indexOf('下載') !== -1 || btns[j].value.indexOf('U') !== -1) {
                                btns[j].click();
                                return true;
                            }
                        }
                    }
                }
                return false;
            """)
            
            if not found:
                print(f"找不到 {ticker} 夾帶簡報檔案的法說會紀錄。")
                return None
                
            print(f"找到 {ticker} 簡報紀錄並已點選下載...")
            
            old_files = set(os.listdir(abs_save_dir))
            
            print("等待法說會簡報 PDF 下載完成...")
            new_filename = wait_for_new_file(abs_save_dir, old_files, timeout=30)
            
            if new_filename:
                old_filepath = os.path.join(abs_save_dir, new_filename)
                new_filepath = os.path.join(abs_save_dir, f"{ticker}_法說會簡報_最新.pdf")
                
                # 如果同名檔案已存在，先刪除
                if os.path.exists(new_filepath):
                    os.remove(new_filepath)
                    
                os.rename(old_filepath, new_filepath)
                print(f"成功下載: {new_filepath}\n")
                return new_filepath
            else:
                print("等待下載逾時。")
                return None
                
        except Exception as table_e:
            print(f"找不到法說會表格或按鈕: {table_e}")
            return None
            
    finally:
        driver.quit()

def _download_mopsov_report_selenium(ticker, url, doc_type, save_dir):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    print(f"啟動瀏覽器準備下載 {ticker} {doc_type}...")
    driver, abs_save_dir = init_driver(save_dir)
    
    try:
        driver.get(url)
        
        # 等待輸入框出現 (尋找真正表單內的 co_id，避開全站搜尋的 keyword)
        co_id_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='text' and @name='co_id']"))
        )
        co_id_input.clear()
        co_id_input.send_keys(str(ticker))
        
        # 點擊查詢按鈕前，我們嘗試填寫年份與季度以確保有資料
        # 以目前 2026 來說，我們預設查詢 114 年 (去年)
        try:
            import datetime
            last_year = str(datetime.datetime.now().year - 1912)
            driver.execute_script(f"""
                var yInputs = document.querySelectorAll('input[name="year"]');
                for (var i=0; i<yInputs.length; i++) {{
                    yInputs[i].value = '{last_year}';
                }}
                var mSelects = document.querySelectorAll('select[name="season"]');
                for (var i=0; i<mSelects.length; i++) {{
                    mSelects[i].value = '04'; // 預設查第四季或年度
                }}
            """)
        except Exception as e:
            print("填寫年份/季度時發生錯誤 (可能該頁面沒有此欄位)，忽略繼續...", e)
        
        # 點擊真正的查詢按鈕 (避開全站搜尋)
        driver.execute_script("""
            var btns = document.querySelectorAll('input[type="button"]');
            var clicked = false;
            for(var i=0; i<btns.length; i++) {
                var oc = btns[i].getAttribute('onclick') || '';
                var nm = btns[i].getAttribute('name') || '';
                if(oc.indexOf('doAction') !== -1 || oc.indexOf('ajax') !== -1 || nm === 'rulesubmit') {
                    btns[i].click();
                    clicked = true;
                    break;
                }
            }
            if (!clicked && btns.length > 0) {
                btns[btns.length-1].click();
            }
        """)
        
        print(f"已送出 {doc_type} 查詢，等待表格渲染...")
        time.sleep(3)
        
        try:
            # 財報跟三書表的頁面通常直接列出該股票最近或所有的檔案，並有一個「檔案下載」或直接就是超連結/按鈕
            # 尋找所有 檔案下載 按鈕
            download_btns = WebDriverWait(driver, 30).until(
                EC.presence_of_all_elements_located((
                    By.XPATH, 
                    "//input[@type='button' and @value='檔案下載']"
                ))
            )
            
            if not download_btns:
                print(f"找不到 {ticker} {doc_type} 紀錄。")
                return None
                
            print(f"找到 {len(download_btns)} 筆 {doc_type} 紀錄，下載最新一筆...")
            
            old_files = set(os.listdir(abs_save_dir))
            
            # 點擊第一筆 (最新的)
            driver.execute_script("arguments[0].click();", download_btns[0])
            
            print(f"等待 {doc_type} PDF 下載完成...")
            new_filename = wait_for_new_file(abs_save_dir, old_files, timeout=30)
            
            if new_filename:
                old_filepath = os.path.join(abs_save_dir, new_filename)
                new_filepath = os.path.join(abs_save_dir, f"{ticker}_{doc_type}_最新.pdf")
                
                # 如果同名檔案已存在，先刪除
                if os.path.exists(new_filepath):
                    os.remove(new_filepath)
                    
                os.rename(old_filepath, new_filepath)
                print(f"成功下載: {new_filepath}\n")
                return new_filepath
            else:
                print("等待下載逾時。")
                return None
                
        except Exception as table_e:
            print(f"找不到 {doc_type} 表格或按鈕: {table_e}")
            return None
            
    finally:
        driver.quit()

def download_financials_selenium(ticker, save_dir):
    """
    從 mopsov 下載最新一期財報 PDF (t57sb01_q1)
    """
    url = "https://mopsov.twse.com.tw/mops/web/t57sb01_q1"
    return _download_mopsov_report_selenium(ticker, url, "財務報告", save_dir)

def download_affiliated_selenium(ticker, save_dir):
    """
    從 mopsov 下載最新關係企業三書表專區 PDF (t57sb01_q10)
    """
    url = "https://mopsov.twse.com.tw/mops/web/t57sb01_q10"
    return _download_mopsov_report_selenium(ticker, url, "關係企業三書表", save_dir)

if __name__ == "__main__":
    download_briefing_selenium("2330", "./test_pack_selenium")
    download_financials_selenium("2330", "./test_pack_selenium")
    download_affiliated_selenium("2330", "./test_pack_selenium")
