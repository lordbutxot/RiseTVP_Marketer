@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"

echo.
echo ====================================
echo   RISE TVP TRADE ROUTE OPTIMIZER
echo ====================================
echo.

echo Select mode:
echo 1. Regular Trade (all opportunities)
echo 2. City-Specific Opportunities
set /p mode="Enter 1 or 2: "

if "%mode%"=="1" (
    "%PYTHON%" "%~dp0ocr_to_excel.py" --mode regular %*
) else if "%mode%"=="2" (
    "%PYTHON%" "%~dp0ocr_to_excel.py" --mode city %*
) else (
    echo Invalid choice.
    pause
    exit /b 1
)

pause
