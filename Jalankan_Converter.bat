@echo off
title OmniConverter - Sistem Konversi Serba Bisa
cd /d "%~dp0"
python main.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Terjadi kendala saat menjalankan OmniConverter.
    pause
)
