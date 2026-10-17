@echo off
cd /d "%~dp0"
echo Installing Stock Tracker...
python -m pip install -r requirements.txt
python -m playwright install chromium
if not exist config.yaml copy config.example.yaml config.yaml
echo.
echo Done! Now open config.yaml in Notepad, add your products and turn on notifications.
pause
