@echo off
REM OneTrack Setup Script for Windows
REM Run this on the other PC to set up OneTrack

echo ========================================
echo    OneTrack Setup Script
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed!
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    echo After installing Python, run this script again.
    pause
    exit /b
)

echo Python found!
python --version
echo.

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo Error installing dependencies!
    pause
    exit /b
)

echo.
echo Dependencies installed successfully!
echo.

REM Get PC's IP address
echo Your PC's IP address:
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4" ^| findstr /v "127.0.0.1"') do set IP=%%a
set IP=%IP: =%
echo %IP%
echo.

echo ========================================
echo    Setup Complete!
echo ========================================
echo.
echo To start OneTrack, run:
echo    python app.py
echo.
echo Then open in browser:
echo    http://%IP%:5000
echo.
echo Sales team can access from same WiFi:
echo    http://%IP%:5000
echo.
echo Press any key to start OneTrack now...
pause >nul

REM Start the app
python app.py
