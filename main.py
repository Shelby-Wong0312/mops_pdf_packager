import argparse
import sys
import os

# 將專案根目錄加入路徑
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from src.utils.downloader import MOPSDownloader

def main():
    parser = argparse.ArgumentParser(description="MOPS PDF Packager for NotebookLM")
    parser.add_argument("ticker", type=str, help="欲抓取的台灣股票代碼 (例如: 2330)")
    parser.add_argument("--year", type=int, default=None, help="指定要抓取的民國年分 (若不指定則自動抓取最新一期)")
    args = parser.parse_args()
    
    downloader = MOPSDownloader(ticker=args.ticker, target_year=args.year)
    downloader.run()

if __name__ == "__main__":
    main()
