@echo off
setlocal
chcp 65001 >nul

echo ==================================================
echo CHUONG TRINH CAI DAT ADD-IN (REAL SHARE)
echo ==================================================

:: 1. Tao file script PowerShell de xu ly Share va Copy duong dan
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
echo # Thu nghiem chia se thu muc (Can quyen Admin nhung neu khong co cung khong sao) >> "%PS_SCRIPT%"
echo $shareName = "HoSoAddinCatalog" >> "%PS_SCRIPT%"
echo net share $shareName"=$installDir" /grant:everyone,read 2>$null >> "%PS_SCRIPT%"
echo. >> "%PS_SCRIPT%"
echo # Tao duong dan UNC xịn >> "%PS_SCRIPT%"
echo $computerName = $env:COMPUTERNAME >> "%PS_SCRIPT%"
echo $uncPath = "\\$computerName\$shareName" >> "%PS_SCRIPT%"
echo. >> "%PS_SCRIPT%"
echo # Neu khong share duoc thi dung duong dan IP noi bo >> "%PS_SCRIPT%"
echo if (!(net share $shareName 2>$null)) { >> "%PS_SCRIPT%"
echo    $drive = $installDir.Substring(0,1) >> "%PS_SCRIPT%"
echo    $pathWithoutDrive = $installDir.Substring(3) >> "%PS_SCRIPT%"
echo    $uncPath = "\\127.0.0.1\$drive`$\$pathWithoutDrive" >> "%PS_SCRIPT%"
echo } >> "%PS_SCRIPT%"
echo. >> "%PS_SCRIPT%"
echo Set-Clipboard -Value $uncPath >> "%PS_SCRIPT%"
echo Start-Process "winword.exe" >> "%PS_SCRIPT%"

:: 2. Chay script
powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
del "%PS_SCRIPT%"

echo.
echo ==================================================
echo QUY TRINH CUOI CUNG:
echo.
echo 1. Toi da COPY duong dan MOI vao Clipboard cho ban.
echo 2. Trong Word, hay XOA (Remove) duong dan cu trong 
echo    "Trusted Add-in Catalogs".
echo 3. Paste (Ctrl + V) duong dan MOI vua copy vao.
echo 4. Bam Add catalog -> Tich Show in Menu -> OK.
echo 5. Tat Word di mo lai.
echo ==================================================
pause
