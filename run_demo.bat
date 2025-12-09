@echo off
title VibeVoice Realtime Demo
echo Starting VibeVoice Realtime Demo...
echo.
echo Server will be available at: http://localhost:3000
echo Press Ctrl+C to stop the server
echo.

set CUDA_VISIBLE_DEVICES=1
cd /d E:\VibeVoice-RT\demo

call C:\Users\james\anaconda3\Scripts\activate.bat
python vibevoice_realtime_demo.py --model_path microsoft/VibeVoice-Realtime-0.5B --port 3000 --device cuda

pause
