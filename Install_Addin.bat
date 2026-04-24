@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

echo ==================================================
echo CHUONG TRINH CAI DAT ADD-IN (FIX UNC PATH)
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
echo # Chuyen doi sang duong dan UNC (\\localhost\c$\...) >> "%PS_SCRIPT%"
echo $drive = $installDir.Substring(0,1) >> "%PS_SCRIPT%"
echo $pathWithoutDrive = $installDir.Substring(3) >> "%PS_SCRIPT%"
echo $uncPath = "\\localhost\$drive`$\$pathWithoutDrive" >> "%PS_SCRIPT%"
echo Set-Clipboard -Value $uncPath >> "%PS_SCRIPT%"
echo. >> "%PS_SCRIPT%"
echo # Mo Word >> "%PS_SCRIPT%"
echo Start-Process "winword.exe" >> "%PS_SCRIPT%"

:: 2. Chay script
powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
del "%PS_SCRIPT%"

echo.
echo ==================================================
echo QUY TRINH FIX LOI CHUOI "HTTPS":
echo.
echo 1. Toi da COPY duong dan dang MANG (UNC) vao Clipboard cho ban.
echo 2. Trong Word, vao: File -^> Options -^> Trust Center -^> Settings.
echo 3. Chon "Trusted Add-in Catalogs".
echo 4. Paste (Ctrl + V) duong dan vao o "Catalog Url". 
echo    (No se co dang: \\localhost\c$\Users\...)
echo 5. Bam "Add Catalog" -> Tich vao "Show in Menu" -> Bam OK.
echo 6. Tat Word di mo lai la se thay tab "HO SO" tren Toolbar.
echo ==================================================
pause
