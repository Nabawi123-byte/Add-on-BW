@echo off
setlocal
cd /d "%~dp0"
title Instalasi Master untuk Word dan Excel
echo.
echo ==============================================
echo   MASTER - asisten perintah Word ^& Excel
echo ==============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [!] Node.js belum terpasang.
  echo     Unduh versi LTS dari https://nodejs.org lalu jalankan file ini lagi.
  pause
  exit /b 1
)

echo [1/4] Memasang komponen (butuh internet, hanya sekali)...
call npm install --no-fund --no-audit
if errorlevel 1 goto gagal

echo [2/4] Membuat sertifikat HTTPS lokal (klik "Yes" bila Windows bertanya)...
call npx office-addin-dev-certs install --days 3650
if errorlevel 1 goto gagal

echo [3/4] Mendaftarkan Master ke Excel dan Word...
call npx office-addin-dev-settings register manifest-excel.xml
call npx office-addin-dev-settings register manifest-word.xml

echo [4/4] Mengatur server Master agar otomatis jalan saat Windows menyala...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut([Environment]::GetFolderPath('Startup')+'\Master Server.lnk'); $s.TargetPath='wscript.exe'; $s.Arguments='\"%~dp0start-master-hidden.vbs\"'; $s.WorkingDirectory='%~dp0'; $s.Save()"
start "" wscript.exe "%~dp0start-master-hidden.vbs"

echo.
echo SELESAI. Tutup lalu buka lagi Excel / Word.
echo Tombol "Master" ada di tab Beranda (Home), di ujung kanan.
echo Jika belum muncul: Beranda ^> Add-in ^> Add-in Saya ^> Add-in Pengembang ^> Master.
echo.
pause
exit /b 0

:gagal
echo.
echo [!] Instalasi gagal. Periksa koneksi internet lalu jalankan lagi.
pause
exit /b 1
