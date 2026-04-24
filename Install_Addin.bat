@echo off
setlocal
chcp 65001 >nul

set "INSTALL_DIR=%APPDATA%\HoSoAddin"
set "GITHUB_URL=https://raw.githubusercontent.com/buiquangtrung2012-ops/HoSoAddin/main"
set "CATALOG_GUID={B1A9908E-1C4F-40E2-9EED-7C919D12DF01}"

echo ==================================================
echo CHUONG TRINH CAI DAT ADD-IN (TU DONG HOAN TOAN)
echo ==================================================

:: 1. Tao thu muc
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo [1/5] Dang tai file manifest.xml moi nhat...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/manifest.xml', '%INSTALL_DIR%\manifest.xml')"

echo [2/5] Dang tai file Mau_ToTrinh.docx...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/Mau_ToTrinh.docx', '%INSTALL_DIR%\Mau_ToTrinh.docx')"

echo [3/5] Don dep cache va cac thiet lap cu...
rmdir /s /q "%LocalAppData%\Microsoft\Office\16.0\Wef" >nul 2>&1
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddin_GitHub" /f >nul 2>&1
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddinDev" /f >nul 2>&1
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /f >nul 2>&1

echo [4/5] Dang ky Trusted Catalog (Add-in Menu)...
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Id" /t REG_SZ /d "%CATALOG_GUID%" /f >nul
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Url" /t REG_SZ /d "%INSTALL_DIR%" /f >nul
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Flags" /t REG_DWORD /d 1 /f >nul

echo [5/5] Dang ky Trusted Location (He thong tin cay)...
set "TRUSTED_LOC_KEY=HKCU\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\HoSoAddin"
REG ADD "%TRUSTED_LOC_KEY%" /v "Path" /t REG_SZ /d "%INSTALL_DIR%" /f >nul
REG ADD "%TRUSTED_LOC_KEY%" /v "AllowSubfolders" /t REG_DWORD /d 1 /f >nul
REG ADD "%TRUSTED_LOC_KEY%" /v "Description" /t REG_SZ /d "HoSoAddin Catalog" /f >nul

echo.
echo ==================================================
echo DA CAI DAT XONG! 
echo.
echo QUY TRINH CUOI CUNG:
echo 1. Tat tat ca cac cua so Word dang mo.
echo 2. Mo lai Word.
echo 3. Vao tab Insert -^> My Add-ins.
echo 4. Bam nut "Refresh" (lam moi) o goc tren ben phai.
echo 5. Tab "SHARED FOLDER" se xuat hien va co Add-in o do.
echo ==================================================
pause
