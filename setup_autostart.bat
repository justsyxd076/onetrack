@echo off
REM OneTrack Auto-Start Setup
REM Run this once to set up auto-start on boot

echo ========================================
echo    OneTrack Auto-Start Setup
echo ========================================
echo.

REM Get the current directory
set APP_DIR=%~dp0
set APP_DIR=%APP_DIR:~0,-1%

echo App directory: %APP_DIR%
echo.

REM Create a VBS script for silent startup
echo Set WshShell = CreateObject("WScript.Shell") > "%APP_DIR%\start_silent.vbs"
echo WshShell.Run "cmd /c cd /d ""%APP_DIR%"" && python app.py", 0, False >> "%APP_DIR%\start_silent.vbs"

REM Create a shortcut in Startup folder
echo Creating startup shortcut...
set STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
copy "%APP_DIR%\start_silent.vbs" "%STARTUP%\OneTrack.vbs" >nul 2>&1

if errorlevel 1 (
    echo Error creating startup shortcut!
    echo Try running as administrator.
    pause
    exit /b
)

echo.
echo ========================================
echo    Setup Complete!
echo ========================================
echo.
echo OneTrack will now:
echo    - Start automatically when Windows boots
echo    - Run silently in background
echo    - Be accessible at http://YOUR-IP:5000
echo.
echo To stop OneTrack:
echo    - Open Task Manager
echo    - Find "python.exe"
echo    - End task
echo.
echo To remove auto-start:
echo    - Delete: %STARTUP%\OneTrack.vbs
echo.
pause
