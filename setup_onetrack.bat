@echo off
REM OneTrack Setup Script for Windows
REM Run this on the other PC to set up OneTrack

echo ========================================
echo    OneTrack Setup Script
echo ========================================
echo.

REM Change to the directory where this script is located
cd /d "%~dp0"

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed!
    echo.
    echo Please install Python from Microsoft Store:
    echo    1. Open Microsoft Store
    echo    2. Search for "Python"
    echo    3. Install Python 3.12
    echo.
    echo After installing Python, run this script again.
    pause
    exit /b
)

echo Python found!
python --version
echo.

REM Check if requirements.txt exists
if not exist "requirements.txt" (
    echo ERROR: requirements.txt not found!
    echo.
    echo Make sure you are running this script from the OneTrack folder.
    echo Current directory: %CD%
    echo.
    pause
    exit /b
)

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
