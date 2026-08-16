@echo off
setlocal enabledelayedexpansion

echo ==============================================
echo 0. Mendeteksi file revisi terbaru...
echo ==============================================

for /f "delims=" %%F in ('dir /b /o:n code_*.gs code_*.js 2^>nul') do set "LATEST_CODE=%%F"
for /f "delims=" %%F in ('dir /b /o:n index_*.html 2^>nul') do set "LATEST_HTML=%%F"

if defined LATEST_CODE (
    echo [Backend] Mengambil versi: !LATEST_CODE!
    copy /y "!LATEST_CODE!" "Kode.js" >nul
)

if defined LATEST_HTML (
    echo [Frontend] Mengambil versi: !LATEST_HTML!
    copy /y "!LATEST_HTML!" "index.html" >nul
)

echo ==============================================
echo 1. Mengunggah backend ke Google Apps Script...
echo ==============================================
call clasp push --force

echo ==============================================
echo 2. Mengunggah frontend ke GitHub...
echo ==============================================
git add .
git commit -m "Update revisi otomatis"
git push origin main

echo ==============================================
echo 3. Membuka website di Google Chrome...
echo ==============================================
start chrome https://gurukemarinsore.github.io/komitesmp3/

echo ==============================================
echo SELESAI! Semua perubahan berhasil di-deploy.
echo ==============================================
pause