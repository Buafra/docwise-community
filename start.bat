@echo off
cd /d "%~dp0"
set "TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe"
set "TESSDATA_PREFIX=%~dp0tessdata"
echo Starting DocWise Community...
echo Open: http://127.0.0.1:8120
echo.
python app.py
pause
