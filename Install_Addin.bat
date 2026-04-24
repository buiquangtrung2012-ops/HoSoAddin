@echo off
setlocal
chcp 65001 >nul

echo ==================================================
echo CHUONG TRINH CAI DAT ADD-IN (OPEN SETTINGS)
echo ==================================================

:: 1. Tao file script PowerShell tam thoi
set "PS_SCRIPT=%TEMP%\install_hoso.ps1"

echo $appdata = [System.Environment]::GetFolderPath('ApplicationData') > "%PS_SCRIPT%"
echo $installDir = Join-Path $appdata "HoSoAddin" >> "%PS_SCRIPT%"
echo if (!(Test-Path $installDir)) { New-Item -ItemType Directory -Path $installDir -Force } >> "%PS_SCRIPT%"
echo. >> "%PS_SCRIPT%"
echo # Tai manifest tu GitHub >> "%PS_SCRIPT%"
echo $url = "https://raw.githubusercontent.com/buiquangtrung2012-ops/HoSoAddin/main/manifest.xml" >> "%PS_SCRIPT%"
echo $dest = Join-Path $installDir "manifest.xml" >> "%PS_SCRIPT%"
echo (New-Object Net.WebClient).DownloadFile($url, $dest) >> "%PS_SCRIPT%"
echo. >> "%PS_SCRIPT%"
echo # Copy duong dan vao Clipboard de nguoi dung chi viec Paste >> "%PS_SCRIPT%"
echo Set-Clipboard -Value $installDir >> "%PS_SCRIPT%"
echo. >> "%PS_SCRIPT%"
echo # Mo Word va tu dong bat cua so Trust Center (neu co the) >> "%PS_SCRIPT%"
echo Start-Process "winword.exe" >> "%PS_SCRIPT%"

:: 2. Chay script
powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
del "%PS_SCRIPT%"

echo.
echo ==================================================
echo QUY TRINH 3 GIAY DE DUC DIEM LOI:
echo.
echo 1. Toi da COPY duong dan thu muc vao Clipboard cho ban.
echo 2. Bay gio trong Word, hay vao: 
echo    File -> Options -> Trust Center -> Trust Center Settings.
echo 3. Chon "Trusted Add-in Catalogs".
echo 4. Click chuot vao o "Catalog Url", nhan CTRL + V roi bam "Add Catalog".
echo 5. Tich vao o "Show in Menu" cua dong vua hien ra.
echo 6. Bam OK, tat Word di mo lai la xong 100%%.
echo ==================================================
pause
