@echo off
rem Rebuilds Stock_Dashboard.html from all monthly stock statement xlsx
rem files in this folder, then opens it in the default browser.
cd /d "%~dp0"
python build_dashboard.py
if errorlevel 1 (
    echo.
    echo Build failed. Make sure Python and openpyxl are installed:
    echo     pip install openpyxl
    pause
    exit /b 1
)
start "" "Stock_Dashboard.html"
