@echo off
goto :MAIN

:MAIN
setlocal
cd /d %~dp0

echo =====================================================
echo WAW WindowAsWallpaper Runner
echo =====================================================
echo.

:: Check for administrator privileges
:: The openfiles command returns an error without admin rights, used here for detection.
openfiles >nul 2>&1
if %errorlevel% neq 0 goto :ADMIN_ERROR

echo Administrator privileges confirmed. Proceeding...

:: Run main script
start .\uv.exe run main.py "settings.json"

:: Closing launcher
echo Closing launcher
timeout /t 2 /nobreak >nul
exit /b 0

:ADMIN_ERROR
echo =====================================================
echo Error: This script requires administrator privileges.
echo Please right-click and select [Run as administrator].
echo =====================================================
pause
exit /b 1