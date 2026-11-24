@echo off
title OmniPDF - iLovePDF Versi Offline
cd /d "%~dp0"
echo ============================================================
echo   Menjalankan OmniPDF (iLovePDF Offline Suite)...
echo   100%% Offline, Aman, Cepat, Tanpa Batas Ukuran File!
echo ============================================================
echo.
python server.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Terjadi kendala saat menjalankan server.
    pause
)
