@echo off
setlocal
chcp 65001 >nul

set INSTALL_DIR=%APPDATA%\HoSoAddin
set GITHUB_URL=https://raw.githubusercontent.com/buiquangtrung2012-ops/HoSoAddin/main

echo ==============================================
echo CHUONG TRINH CAI DAT ADD-IN
echo ==============================================

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo [1/3] Dang tai file manifest.xml...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/manifest.xml', '%INSTALL_DIR%\manifest.xml')"

echo [2/3] Dang tai file Mau_ToTrinh.docx...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/Mau_ToTrinh.docx', '%INSTALL_DIR%\Mau_ToTrinh.docx')"

echo [3/3] Dang dang ky voi Microsoft Word...
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddin_GitHub" /t REG_SZ /d "%INSTALL_DIR%\manifest.xml" /f >nul

echo ==============================================
echo DA HOAN THANH! DANG MO FILE MAU...
echo ==============================================
if exist "%INSTALL_DIR%\Mau_ToTrinh.docx" (
    start "" "%INSTALL_DIR%\Mau_ToTrinh.docx"
) else (
    echo [LOI] Khong tim thay file Mau_ToTrinh.docx sau khi tai.
)
pause
