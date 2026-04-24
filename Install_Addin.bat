@echo off
setlocal
chcp 65001 >nul

set "INSTALL_DIR=%APPDATA%\HoSoAddin"
set "GITHUB_URL=https://raw.githubusercontent.com/buiquangtrung2012-ops/HoSoAddin/main"

echo ==================================================
echo CHUONG TRINH CAI DAT ADD-IN (FIX TRUNCATION)
echo ==================================================

:: 1. Tao thu muc
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo [1/4] Dang tai file manifest.xml moi nhat...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/manifest.xml', '%INSTALL_DIR%\manifest.xml')"

echo [2/4] Dang tai file Mau_ToTrinh.docx...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/Mau_ToTrinh.docx', '%INSTALL_DIR%\Mau_ToTrinh.docx')"

echo [3/4] Xoa bo nho dem va cac thiet lap cu...
rmdir /s /q "%LocalAppData%\Microsoft\Office\16.0\Wef" >nul 2>&1

echo [4/4] Dang ky he thong tin cay bang PowerShell (Cuc ky chinh xac)...
powershell -ExecutionPolicy Bypass -Command ^
    "$dir = '%INSTALL_DIR%';" ^
    "$guid = '{B1A9908E-1C4F-40E2-9EED-7C919D12DF01}';" ^
    "$wefPath = \"HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$guid\";" ^
    "if (!(Test-Path $wefPath)) { New-Item -Path $wefPath -Force };" ^
    "Set-ItemProperty -Path $wefPath -Name 'Id' -Value $guid;" ^
    "Set-ItemProperty -Path $wefPath -Name 'Url' -Value $dir;" ^
    "Set-ItemProperty -Path $wefPath -Name 'Flags' -Value 1 -PropertyType DWord;" ^
    "$wordLoc = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\HoSoAddin';" ^
    "if (!(Test-Path $wordLoc)) { New-Item -Path $wordLoc -Force };" ^
    "Set-ItemProperty -Path $wordLoc -Name 'Path' -Value $dir;" ^
    "Set-ItemProperty -Path $wordLoc -Name 'AllowSubfolders' -Value 1 -PropertyType DWord;" ^
    "Set-ItemProperty -Path $wordLoc -Name 'Description' -Value 'HoSoAddin Catalog';"

echo.
echo ==================================================
echo DA CAI DAT XONG! 
echo.
echo QUY TRINH CUOI CUNG:
echo 1. Tat tat ca cac cua so Word dang mo.
echo 2. Mo lai Word.
echo 3. Vao tab Insert -^> My Add-ins.
echo 4. Bam nut "Refresh" (lam moi) o goc tren ben phai.
echo 5. Tab "SHARED FOLDER" se hien ra va co Add-in o do.
echo ==================================================
pause
