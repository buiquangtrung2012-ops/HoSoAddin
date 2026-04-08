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
echo   QUET SACH TAN DU VA DANG KY BAN MOI (NEW)
echo ==============================================

:: 1. Dong Word (Phai dong Word moi xoa duoc Cache)
echo Dang dong Word...
taskkill /F /IM winword.exe >nul 2>&1

:: 2. Xoa Cache sau cua Office
echo Dang xoa bo nho dem Office (Wef)...
if exist "%localappdata%\Microsoft\Office\16.0\Wef" (
    rmdir /s /q "%localappdata%\Microsoft\Office\16.0\Wef" >nul 2>&1
)

:: 3. Xoa sach Registry de lam moi hoan toan
echo Dang don dep Registry...
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs" /f >nul 2>&1
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /f >nul 2>&1

:: 4. Xoa bỏ hoàn toàn các Network Share
echo Dang go bo chia se mang...
net share HoSoAddin /delete >nul 2>&1
net share HoSoAddinLocal /delete >nul 2>&1

:: 5. Dang ky lai voi ma ID moi (F3E2...)
echo Dang dang ky ban moi Quan Ly Ho So (NEW)...
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddinDevNew" /t REG_SZ /d "%~dp0manifest.xml" /f >nul

echo.
echo ==============================================
echo HOAN THANH DON DEP! 
echo 1. Ban hay mo lai Word.
echo 2. Vao Insert -> My Add-ins.
echo 3. Nhấn vào bản MOI tên là: "Quan Ly Ho So (NEW)".
echo ==============================================
pause
