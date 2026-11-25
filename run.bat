@echo off
title Lab Staff Feedback Form - KEEP THIS WINDOW OPEN
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python from https://www.python.org and tick "Add python.exe to PATH".
  pause
  exit /b 1
)

python -c "import flask, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Installing required packages...
  python -m pip install -r requirements.txt
)

rem Open the browser a few seconds after the server starts
start "" /b cmd /c "timeout /t 4 /nobreak >nul & start http://127.0.0.1:5000"

echo.
echo  Starting the feedback form... keep this window open while using the site.
echo  To stop it, close this window or press Ctrl+C.
echo.
python app.py
pause
