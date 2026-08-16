@echo off
echo ==============================================
echo 1. Mengunggah backend ke Google Apps Script...
echo ==============================================
call clasp push --force

echo ==============================================
echo 2. Mengunggah frontend ke GitHub...
echo ==============================================
git add .
git commit -m "Update otomatis"
git push origin main

echo ==============================================
echo 3. Membuka website di Google Chrome...
echo ==============================================
start chrome https://gurukemarinsore.github.io/komitesmp3/

echo ==============================================
echo SELESAI! Semua perubahan berhasil di-deploy.
echo ==============================================
pause