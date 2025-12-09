@echo off
title VibeVoice SAPI Pipe Server
echo ============================================
echo   VibeVoice SAPI Named Pipe Server
echo ============================================
echo.
echo This server provides TTS through a named pipe
echo for the SAPI DLL to connect to.
echo.
echo Pipe: \\.\pipe\vibevoice
echo.
echo Press Ctrl+C to stop the server
echo.

cd /d "%~dp0demo"

python sapi_pipe_server.py --model_path microsoft/VibeVoice-Realtime-0.5B --device cuda:0

pause
