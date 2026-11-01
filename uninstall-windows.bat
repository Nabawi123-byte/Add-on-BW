@echo off
cd /d "%~dp0"
echo Menghapus Master dari Excel dan Word...
call npx office-addin-dev-settings unregister manifest-excel.xml
call npx office-addin-dev-settings unregister manifest-word.xml
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Master Server.lnk" 2>nul
taskkill /F /FI "WINDOWTITLE eq Server Master" >nul 2>nul
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":3000" ^| findstr LISTENING') do taskkill /F /PID %%p >nul 2>nul
echo Selesai. Folder ini boleh dihapus.
pause
