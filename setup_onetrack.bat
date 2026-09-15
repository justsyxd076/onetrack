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
echo Choose how to run OneTrack:
echo.
echo    1. Normal mode (manual start)
echo    2. Watchdog mode (auto-restart if crashed)
echo    3. Auto-start on boot (recommended)
echo.
set /p choice="Enter choice (1-3): "

if "%choice%"=="1" (
    echo.
    echo Starting OneTrack...
    echo Press Ctrl+C to stop.
    echo.
    python app.py
) else if "%choice%"=="2" (
    echo.
    echo Starting OneTrack with watchdog...
    echo App will auto-restart if it crashes.
    echo Press Ctrl+C to stop.
    echo.
    python watchdog.py
) else if "%choice%"=="3" (
    echo.
    echo Setting up auto-start on boot...
    call setup_autostart.bat
    echo.
    echo Starting OneTrack with watchdog...
    python watchdog.py
) else (
    echo.
    echo Invalid choice. Starting in normal mode...
    python app.py
)
