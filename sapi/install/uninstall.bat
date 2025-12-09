@echo off
setlocal enabledelayedexpansion

:: VibeVoice SAPI Uninstallation Script
:: Run this script as Administrator

echo ============================================
echo   VibeVoice SAPI TTS Uninstallation
echo ============================================
echo.

:: Check for admin privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This script must be run as Administrator.
    echo Right-click and select "Run as administrator"
    pause
    exit /b 1
)

:: Get script directory
set "SCRIPT_DIR=%~dp0"
set "SAPI_DIR=%SCRIPT_DIR%.."
set "PROJECT_DIR=%SAPI_DIR%\.."
set "DLL_PATH=%SAPI_DIR%\VibeVoiceSAPI\bin\Release\VibeVoiceSAPI.dll"

echo Step 1: Stopping service...
net stop VibeVoiceTTS >nul 2>&1
echo   Done.

echo.
echo Step 2: Removing service...
cd /d "%PROJECT_DIR%\service"
python vibevoice_service.py remove >nul 2>&1
echo   Done.

echo.
echo Step 3: Unregistering voice tokens...
powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%register_voices.ps1" -Uninstall

echo.
echo Step 4: Unregistering COM DLL...
if exist "%DLL_PATH%" (
    regsvr32 /u /s "%DLL_PATH%"
    echo   Done.
) else (
    echo   DLL not found, skipping.
)

echo.
echo ============================================
echo   Uninstallation Complete!
echo ============================================
echo.
pause
