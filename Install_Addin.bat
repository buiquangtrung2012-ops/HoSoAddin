@echo off
setlocal
chcp 65001 >nul

set INSTALL_DIR=%APPDATA%\HoSoAddin
set GITHUB_URL=https://raw.githubusercontent.com/buiquangtrung2012-ops/HoSoAddin/main
set CATALOG_GUID={B1A9908E-1C4F-40E2-9EED-7C919D12DF01}

echo ==============================================
echo CHUONG TRINH CAI DAT ADD-IN (TRUSTED CATALOG)
echo ==============================================

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo [1/4] Dang tai file manifest.xml...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/manifest.xml', '%INSTALL_DIR%\manifest.xml')"

echo [2/4] Dang tai file Mau_ToTrinh.docx...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/Mau_ToTrinh.docx', '%INSTALL_DIR%\Mau_ToTrinh.docx')"

echo [3/4] Xoa ban Developer cu (neu co) de tranh loi xung dot...
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddin_GitHub" /f >nul 2>&1
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddinDev" /f >nul 2>&1

echo [4/4] Dang dang ky thu muc Add-in vao Trust Center...
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Id" /t REG_SZ /d "%CATALOG_GUID%" /f >nul
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Url" /t REG_SZ /d "%INSTALL_DIR%" /f >nul
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Flags" /t REG_DWORD /d 1 /f >nul

echo ==============================================
echo DA HOAN THANH CAI DAT ADD-IN CHINH THUC! 
echo.
echo ==============================================
echo LUU Y QUAN TRONG DE KHONG BAO GIO BI LOI:
echo 1. Hay mo mot file Word TRONG (Blank Document).
echo 2. Vao tab Insert -^> My Add-ins -^> Chon tab SHARED FOLDER.
echo 3. Bam vao "Quan Ly Ho So" de mo Add-in lan dau tien.
echo 4. Chinh sua form sau do LUU LAI thanh "Mau_ToTrinh.docx" moi nhat.
echo 5. Tu gio tro di, viec luu file se tuyet doi an toan.
echo ==============================================
pause
