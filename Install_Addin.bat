@echo off
title HoSo Addin Installer
color 0A

set "INSTALL_DIR=%APPDATA%\HoSoAddin"
set "MANIFEST_PATH=%INSTALL_DIR%\manifest.xml"
set "GITHUB_RAW_URL=https://raw.githubusercontent.com/buiquangtrung2012-ops/HoSoAddin/main/manifest.xml"

echo Installing HoSo Add-in...
echo.

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo Downloading manifest...
powershell -Command "Invoke-WebRequest -Uri '%GITHUB_RAW_URL%' -OutFile '%MANIFEST_PATH%'"

if not exist "%MANIFEST_PATH%" (
    echo Download failed.
    pause
    exit /b
)

echo Closing Word...
taskkill /f /im WINWORD.EXE >nul 2>&1

echo Cleaning Office cache...
rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1
rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\WebServiceCache" >nul 2>&1

echo Registering Developer Add-in...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" ^
 /v "HoSoAddin" ^
 /t REG_SZ ^
 /d "%MANIFEST_PATH%" ^
 /f

echo Opening Word...
start winword

echo.
echo Done!
pause