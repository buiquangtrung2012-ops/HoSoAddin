@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

REM Check Admin Rights
openfiles >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] VUI LONG CHAY FILE NAY BANG QUYEN ADMINISTRATOR!
    pause
    exit /b
)

echo ==============================================
echo DANG THIET LAP BAN GITHUB (FINAL FIX)
echo ==============================================

:: 1. Version tag
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set datetime=%%I
set "v_tag=%datetime:~6,2%%datetime:~4,2%%datetime:~0,4%.%datetime:~8,4%"

:: 2. Update manifest.xml
powershell -Command "(Get-Content manifest.xml) -replace '<SourceLocation DefaultValue=\".*?\" />', '<SourceLocation DefaultValue=\"https://buiquangtrung2012-ops.github.io/HoSoAddin/taskpane.html\" />' | Set-Content manifest.xml"

:: 3. Update taskpane.html
powershell -Command "(Get-Content taskpane.html) -replace '<div id=\"statusInfo\" class=\".*?\">.*?</div>', '<div id=\"statusInfo\" class=\"text-[10px] font-bold px-2 py-1 bg-slate-100 rounded-full text-slate-500 uppercase tracking-widest\">v%v_tag%</div>' | Set-Content taskpane.html"
powershell -Command "(Get-Content taskpane.html) -replace '<script type=\"module\" src=\"taskpane.js.*?\">', '<script type=\"module\" src=\"taskpane.js?v=%v_tag%\">' | Set-Content taskpane.html"

:: 4. Clean up old shares
net share HoSoAddinLocal /delete >nul 2>&1
net share HoSoAddin /delete >nul 2>&1
net share WordHoSoAddin /delete >nul 2>&1

:: 5. Register new ID to Registry
echo Dang dang ky vao Registry...
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddinDev" /t REG_SZ /d "%~dp0manifest.xml" /f >nul

echo.
echo ==============================================
echo DA THIET LAP XONG! 
echo 1. Hay mo Word - Insert - My Add-ins.
echo 2. Neu thay ban cu bi loi, hay chuot phai - Remove from list.
echo 3. Chon ban moi "Quan Ly Ho So" trong tab MY ADD-INS.
echo ==============================================
pause
