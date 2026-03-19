@echo off
chcp 65001 >nul 2>&1
title 打包 MOPS PDF Packager 為 EXE

echo ============================================
echo   打包 MOPS PDF Packager 為 EXE
echo ============================================
echo.

REM 切到 bat 所在目錄
cd /d "%~dp0"
REM TODO: gui.py 已移除，改為打包 batch_download.py 或 main.py
if not exist "batch_download.py" (
    echo [錯誤] 找不到 batch_download.py! 請在 mops_pdf_packager 資料夾內執行此檔案。
    pause
    exit /b 1
)

REM 確認 Python 可用
python --version >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 Python! 請先安裝 Python 3.8+
    pause
    exit /b 1
)

REM 用 python -m pip (相容 Microsoft Store 版 Python)
echo [1/3] 檢查並安裝 PyInstaller...
python -m pip install pyinstaller --quiet
if errorlevel 1 (
    echo [錯誤] 安裝 PyInstaller 失敗!
    pause
    exit /b 1
)

echo [2/3] 安裝相依套件...
python -m pip install -r requirements.txt --quiet

REM 打包
echo [3/3] 正在打包為 EXE (可能需要 1~2 分鐘)...
echo.

python -m PyInstaller --noconfirm --onefile --windowed ^
    --name "MOPS_PDF_Packager" ^
    --add-data "src;src" ^
    --hidden-import "src" ^
    --hidden-import "src.scrapers" ^
    --hidden-import "src.scrapers.ebook_scraper" ^
    --hidden-import "src.scrapers.briefing_scraper" ^
    --hidden-import "src.scrapers.esg_scraper" ^
    --hidden-import "src.scrapers.mopsov_scraper" ^
    --hidden-import "src.utils" ^
    --hidden-import "src.utils.downloader" ^
    --hidden-import "requests" ^
    --hidden-import "bs4" ^
    --hidden-import "urllib3" ^
    batch_download.py

echo.
if exist "dist\MOPS_PDF_Packager.exe" (
    echo ============================================
    echo   打包成功!
    echo   EXE 檔案位置: dist\MOPS_PDF_Packager.exe
    echo ============================================
    echo.
    echo 你可以把 dist\MOPS_PDF_Packager.exe 複製到任何地方使用，
    echo 也可以直接上傳到 GitHub Releases。
    echo.

    REM 自動複製到專案根目錄方便取用
    copy "dist\MOPS_PDF_Packager.exe" "MOPS_PDF_Packager.exe" >nul 2>&1

    explorer dist
) else (
    echo [錯誤] 打包失敗! 請檢查上面的錯誤訊息。
)

echo.
pause
