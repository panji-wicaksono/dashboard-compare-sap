@echo off
cd /d "%~dp0"
echo Starting GR Compare Dashboard...
echo Buka browser ke: http://localhost:8000
echo Tekan Ctrl+C untuk menghentikan server.
echo.
python app.py
pause
