@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found! Please install Python first.
    pause
    exit /b 1
)

python push_github.py
pause
