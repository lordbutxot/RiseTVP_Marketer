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
echo 3. Trade Route (multi-hop chained routes)
set /p mode="Enter 1, 2 or 3: "

if "%mode%"=="1" (
    "%PYTHON%" "%~dp0ocr_to_excel.py" --mode regular %*
) else if "%mode%"=="2" (
    "%PYTHON%" "%~dp0ocr_to_excel.py" --mode city %*
) else if "%mode%"=="3" (
    "%PYTHON%" "%~dp0ocr_to_excel.py" --mode route %*
) else (
    echo Invalid choice.
    pause
    exit /b 1
)

pause
