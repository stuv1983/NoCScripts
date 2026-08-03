@echo off
cd /d "%~dp0"
py -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo openpyxl is not installed.
    echo.
    echo Run this command first:
    echo     py -m pip install openpyxl
    echo.
    pause
    exit /b 1
)
py "%~dp0patch_status_checker.py"
if errorlevel 1 pause
