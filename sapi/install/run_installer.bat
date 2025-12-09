@echo off
:: VibeVoice Installer Launcher
:: This will prompt for Administrator privileges

echo Launching VibeVoice Installer...

:: Check if already running as admin
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process '%~dp0vibevoice_installer.py' -Verb RunAs"
) else (
    python "%~dp0vibevoice_installer.py"
)
