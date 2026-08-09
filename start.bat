@echo off
setlocal
cd /d "%~dp0"
set "TESSDATA_PREFIX=%~dp0tessdata"
set "PY=%~dp0.venv\Scripts\python.exe"

if not exist "%PY%" goto :setup
"%PY%" -c "import fastapi, uvicorn" >nul 2>nul
if errorlevel 1 goto :setup
goto :run

:setup
echo ============================================================
echo  DocWise Community - first run setup
echo  Installs Python, Tesseract OCR and app dependencies
echo  automatically. Takes a few minutes. Needs internet.
echo ============================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1"
if errorlevel 1 goto :setupfailed
if not exist "%PY%" goto :setupfailed
"%PY%" -c "import fastapi, uvicorn" >nul 2>nul
if errorlevel 1 goto :setupfailed

:run
echo.
echo Starting DocWise Community...
echo Open: http://127.0.0.1:8120
echo.
"%PY%" app.py
pause
exit /b 0

:setupfailed
echo.
echo Setup did not finish. Read the messages above, fix the issue,
echo then run start.bat again.
echo If Python could not be installed automatically, get it from:
echo   https://www.python.org/downloads/  (tick "Add python.exe to PATH")
echo.
pause
exit /b 1
