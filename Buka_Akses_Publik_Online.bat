@echo off
title OmniPDF - Publikasikan ke Internet (Cloudflare Tunnel)
cd /d "%~dp0"

echo =======================================================================
echo   OmniPDF - Akses Publik Online Tanpa Perlu Akun / Biaya Hosting
echo =======================================================================
echo.
echo 1. Menjalankan server OmniPDF lokal...
start "OmniPDF Server" /min cmd /c "python server.py"

echo 2. Menghubungkan ke Cloudflare Tunnel untuk mendapatkan link HTTPS publik...
echo    Tunggu beberapa detik hingga URL publik (https://...trycloudflare.com) muncul di bawah ini:
echo -----------------------------------------------------------------------
cloudflared tunnel --url http://127.0.0.1:5000
pause
