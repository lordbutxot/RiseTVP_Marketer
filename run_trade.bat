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

"%PYTHON%" "%~dp0ocr_to_excel.py" %*

pause
