@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

git --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Git not found! Please install: https://git-scm.com/
    pause
    exit /b 1
)

git rev-parse --git-dir >nul 2>&1
if %errorlevel% neq 0 (
    git clone https://github.com/Shelby-Wong0312/mops_pdf_packager.git .
    pause
    exit /b 0
)

git pull
pause
