@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

set "INSTALL_DIR=%APPDATA%\HoSoAddin"
set "GITHUB_URL=https://raw.githubusercontent.com/buiquangtrung2012-ops/HoSoAddin/main"
set "CATALOG_GUID={B1A9908E-1C4F-40E2-9EED-7C919D12DF01}"

echo ==================================================
echo CHUONG TRINH CAI DAT ADD-IN (FIX SHARED FOLDER)
echo ==================================================

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo [1/5] Dang tai file manifest.xml...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/manifest.xml', '%INSTALL_DIR%\manifest.xml')"

echo [2/5] Dang tai file Mau_ToTrinh.docx...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/Mau_ToTrinh.docx', '%INSTALL_DIR%\Mau_ToTrinh.docx')"

echo [3/5] Xoa cache Office de cap nhat cau hinh moi...
rmdir /s /q "%LocalAppData%\Microsoft\Office\16.0\Wef" >nul 2>&1

echo [4/5] Don dep cac dang ky cu...
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddin_GitHub" /f >nul 2>&1
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddinDev" /f >nul 2>&1
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /f >nul 2>&1

echo [5/5] Dang dang ky Trusted Catalog (UNC Path)...
:: Chuyen C:\Users\... thanh \\localhost\c$\Users\...
set "DRIVE=%INSTALL_DIR:~0,1%"
set "FOLDER_PATH=%INSTALL_DIR:~3%"
set "UNC_PATH=\\localhost\!DRIVE!$\!FOLDER_PATH!"

REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Id" /t REG_SZ /d "%CATALOG_GUID%" /f >nul
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Url" /t REG_SZ /d "!UNC_PATH!" /f >nul
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Flags" /t REG_DWORD /d 1 /f >nul

echo.
echo ==================================================
echo DA CAI DAT XONG! QUY TRINH MO ADD-IN:
echo 1. Hay MO LAI Word (neu dang mo thi tat di mo lai).
echo 2. Vao tab Insert -> My Add-ins.
echo 3. Bam nut "Refresh" (lam moi) o goc tren ben phai cua bang hien ra.
echo 4. Luc nay tab "SHARED FOLDER" se xuat hien canh tab "MY ADD-INS".
echo 5. Chon "Quan Ly Ho So" va bam Add.
echo ==================================================
pause
