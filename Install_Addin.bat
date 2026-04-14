@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

set MANIFEST_NAME=manifest.xml
set INSTALL_DIR=%APPDATA%\HoSoAddin
set MANIFEST_PATH=%INSTALL_DIR%\%MANIFEST_NAME%
set GITHUB_RAW_URL=https://raw.githubusercontent.com/buiquangtrung2012-ops/HoSoAddin/main/manifest.xml

echo ==============================================
echo CHUONG TRINH CAI DAT ADD-IN
echo ==============================================

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo [1/2] Dang tai file tu GitHub...
powershell -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_RAW_URL%', '%MANIFEST_PATH%')"

echo [2/2] Dang dang ky voi Word...
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddin_GitHub" /t REG_SZ /d "%MANIFEST_PATH%" /f >nul

echo ==============================================
echo DA HOAN THANH!
echo ==============================================
pause
