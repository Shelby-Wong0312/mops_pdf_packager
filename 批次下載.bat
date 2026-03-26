@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found! Please install Python first.
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

python -m pip install -r requirements.txt -q >nul 2>&1
python src\batch_download.py
pause
