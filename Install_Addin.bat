@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

set MANIFEST_NAME=manifest.xml
set DOC_NAME=Mau_ToTrinh.docx
set INSTALL_DIR=%APPDATA%\HoSoAddin
set MANIFEST_PATH=%INSTALL_DIR%\%MANIFEST_NAME%
set DOC_PATH=%INSTALL_DIR%\%DOC_NAME%

set GITHUB_BASE_URL=https://raw.githubusercontent.com/buiquangtrung2012-ops/HoSoAddin/main
set MANIFEST_URL=%GITHUB_BASE_URL%/manifest.xml
set DOC_URL=%GITHUB_BASE_URL%/Mau_ToTrinh.docx

echo ==============================================
echo CHUONG TRINH CAI DAT ADD-IN
echo ==============================================

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo [1/3] Dang tai file manifest...
powershell -Command "(New-Object Net.WebClient).DownloadFile('%MANIFEST_URL%', '%MANIFEST_PATH%')"

echo [2/3] Dang tai file mẫu Mau_ToTrinh.docx...
powershell -Command "(New-Object Net.WebClient).DownloadFile('%DOC_URL%', '%DOC_PATH%')"

echo [3/3] Dang dang ky voi Word...
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddin_GitHub" /t REG_SZ /d "%MANIFEST_PATH%" /f >nul

echo ==============================================
echo DA HOAN THANH! DANG MO FILE MAU...
echo ==============================================
start "" "%DOC_PATH%"
pause
