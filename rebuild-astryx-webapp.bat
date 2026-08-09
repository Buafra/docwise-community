@echo off
cd /d "%~dp0web-astryx"
echo Building DocWise Community Astryx web app...
npm install
npm run build
cd /d "%~dp0"
if exist static\manual.html copy /Y static\manual.html static_astryx\manual.html >nul
if exist static\delete-help.html copy /Y static\delete-help.html static_astryx\delete-help.html >nul
echo Done. Restart start.bat and open http://127.0.0.1:8120
pause
