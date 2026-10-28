@echo off
title Bigrock Bulk Mailer - Web Dashboard
color 0B
echo =======================================================
echo          Bigrock Bulk Mailer Web Interface
echo =======================================================
echo.

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not found in your system PATH.
    echo Please install Python 3.8+ and check "Add Python to PATH".
    pause
    exit /b 1
)

echo [1/2] Verifying Python libraries...
python -m pip install -r requirements.txt --quiet

echo [2/2] Starting local web application server...
echo Your default web browser will open automatically at http://127.0.0.1:8000
echo (Keep this black command window open while using the mailer)
echo.

python app.py

echo.
echo Server stopped.
pause
