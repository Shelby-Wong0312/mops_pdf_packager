@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

echo ============================================================
echo   MOPS PDF Packager - 批次下載
echo ============================================================
echo.

:: 檢查 Python 是否已安裝
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [錯誤] 找不到 Python！
    echo.
    echo 請先安裝 Python:
    echo   1. 前往 https://www.python.org/downloads/
    echo   2. 下載並安裝 Python 3.8 以上版本
    echo   3. 安裝時請勾選「Add Python to PATH」
    echo.
    pause
    exit /b 1
)

:: 安裝依賴套件（靜默模式）
echo 正在檢查並安裝依賴套件...
python -m pip install -r requirements.txt -q >nul 2>&1
echo 依賴套件已就緒。
echo.

:: 檢查公司清單是否存在
if not exist "公司清單.xlsx" (
    echo [錯誤] 找不到「公司清單.xlsx」！
    echo.
    echo 請先在專案目錄下建立「公司清單.xlsx」，格式如下:
    echo   A欄: 股票代碼（必填，例如 2330）
    echo   B欄: 起始年份（民國年，選填）
    echo   C欄: 結束年份（民國年，選填）
    echo.
    pause
    exit /b 1
)

:: 執行批次下載
echo 開始批次下載...
echo.
python batch_download.py

echo.
echo ============================================================
echo   批次下載執行完畢
echo ============================================================
echo.
pause
