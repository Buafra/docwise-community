@echo off
setlocal
cd /d "%~dp0"
set "TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe"
set "TESSDATA_PREFIX=%~dp0tessdata"
set "PY=%~dp0.venv\Scripts\python.exe"

if exist "%PY%" goto :deps

echo First run: creating Python virtual environment...
where py >nul 2>nul
if errorlevel 1 goto :trypython
py -3 -m venv .venv
goto :checkvenv

:trypython
python -m venv .venv

:checkvenv
if not exist "%PY%" goto :nopython

:deps
"%PY%" -c "import fastapi, uvicorn" >nul 2>nul
if not errorlevel 1 goto :run
echo Installing dependencies (first run only, needs internet)...
"%PY%" -m pip install --upgrade pip
"%PY%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: Failed to install dependencies.
    echo Check your internet connection and run start.bat again.
    pause
    exit /b 1
)

:run
echo.
echo Starting DocWise Community...
echo Open: http://127.0.0.1:8120
echo.
"%PY%" app.py
pause
exit /b 0

:nopython
echo.
echo ERROR: Python was not found on this computer.
echo.
echo Install Python 3.10 or newer from:
echo   https://www.python.org/downloads/
echo.
echo IMPORTANT: tick "Add python.exe to PATH" during installation,
echo then run start.bat again.
echo.
pause
exit /b 1
