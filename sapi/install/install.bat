@echo off
setlocal enabledelayedexpansion

:: VibeVoice SAPI Installation Script
:: Run this script as Administrator

echo ============================================
echo   VibeVoice SAPI TTS Installation
echo ============================================
echo.

:: 1. Check for admin privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This script must be run as Administrator.
    echo Right-click and select "Run as administrator"
    pause
    exit /b 1
)

:: 2. Setup Paths (Normalized)
pushd "%~dp0"
set "SCRIPT_DIR=%CD%"
cd ..
set "ROOT_DIR=%CD%"
popd

:: ADJUST THIS PATH based on where Visual Studio built your DLL
:: Standard VS 2022 path is usually x64\Release
set "DLL_PATH=%ROOT_DIR%\VibeVoiceSAPI\bin\Release"

:: 3. Check if DLL exists
if not exist "%DLL_PATH%" (
    echo ERROR: VibeVoiceSAPI.dll not found at:
    echo   "%DLL_PATH%"
    echo.
    echo Please build the solution first:
    echo   1. Open VibeVoiceSAPI.sln in Visual Studio
    echo   2. Select "Release" and "x64" in the toolbar
    echo   3. Build Solution (Ctrl+Shift+B)
    echo.
    pause
    exit /b 1
)

:: 4. Register the COM DLL
echo Step 1: Registering COM DLL...
regsvr32 /s "%DLL_PATH%"
if %errorLevel% neq 0 (
    echo ERROR: Failed to register DLL. 
    echo ensure you have the Visual C++ Redistributables installed.
    pause
    exit /b 1
)
echo   Success.

:: 5. Register Voices (Using the Python script)
echo.
echo Step 2: Registering voice tokens...
:: We use the Python script from the previous step here
python "%SCRIPT_DIR%\register_voice.py"
if %errorLevel% neq 0 (
    echo ERROR: Failed to register voices.
    echo Ensure Python is in your PATH and dependencies are installed.
    pause
    exit /b 1
)

:: 6. Optional: Service Setup
:: Only runs if you have a service script ready
set "SERVICE_SCRIPT=%ROOT_DIR%\service\vibevoice_service.py"
if exist "%SERVICE_SCRIPT%" (
    echo.
    echo Step 3: Installing Python service...
    python "%SERVICE_SCRIPT%" install
    if !errorLevel! neq 0 (
        echo WARNING: Service installation failed. Run manually if needed.
    )
) else (
    echo.
    echo Step 3: Service script not found, skipping service installation.
    echo (Expected at: %SERVICE_SCRIPT%)
)

echo.
echo ============================================
echo   Installation Complete!
echo ============================================
echo.
echo To test:
echo   1. Go to Windows Settings - Time ^& Language - Speech
echo   2. Select "VibeVoice - Carter"
echo   3. Click "Preview Voice"
echo.
echo NOTE: Ensure server.py is running!
echo.
pause