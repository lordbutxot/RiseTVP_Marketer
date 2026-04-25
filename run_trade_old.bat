@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%PYTHON%" (
    set "PYTHON=python"
)

REM Preguntar por presupuesto
echo.
echo ====================================
echo   RISE TVP TRADE ROUTE OPTIMIZER
echo ====================================
echo.
echo ¿Deseas usar un presupuesto inicial?
echo (Déjalo en blanco para omitir)
set /p BUDGET="Ingresa el presupuesto en CR: "

REM Construir comando con o sin presupuesto
if "%BUDGET%"=="" (
    "%PYTHON%" "%~dp0ocr_to_excel.py" %*
) else (
    "%PYTHON%" "%~dp0ocr_to_excel.py" --budget %BUDGET% %*
)

pause
