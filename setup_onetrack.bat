@echo off
REM OneTrack Setup Script for Windows

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
    echo Please install Python from Microsoft Store first.
    echo.
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
    echo Current directory: %CD%
    echo.
    echo Make sure you opened the correct folder.
    echo.
    pause
    exit /b
)

echo Found requirements.txt
echo.

REM Install dependencies
echo Installing dependencies... This may take a minute.
echo.
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Error installing dependencies!
    echo.
    pause
    exit /b
)

echo.
echo ========================================
echo    Dependencies Installed!
echo ========================================
echo.

REM Get PC's IP address
echo Your PC's IP address:
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4" ^| findstr /v "127.0.0.1"') do set IP=%%a
set IP=%IP: =%
echo %IP%
echo.

echo ========================================
echo    Choose How to Run OneTrack
echo ========================================
echo.
echo    1. Normal mode (manual start)
echo    2. Watchdog mode (auto-restart)
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
    echo Starting in normal mode...
    python app.py
)

pause
