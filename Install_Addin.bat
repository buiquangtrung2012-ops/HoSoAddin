@echo off
setlocal
chcp 65001 >nul

set "INSTALL_DIR=%APPDATA%\HoSoAddin"
set "GITHUB_URL=https://raw.githubusercontent.com/buiquangtrung2012-ops/HoSoAddin/main"
set "CATALOG_GUID={B1A9908E-1C4F-40E2-9EED-7C919D12DF01}"

echo ==================================================
echo CHUONG TRINH CAI DAT ADD-IN (FINAL STABLE)
echo ==================================================

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo [1/4] Dang tai file manifest.xml moi nhat...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/manifest.xml', '%INSTALL_DIR%\manifest.xml')"

echo [2/4] Dang tai file Mau_ToTrinh.docx...
powershell -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('%GITHUB_URL%/Mau_ToTrinh.docx', '%INSTALL_DIR%\Mau_ToTrinh.docx')"

echo [3/4] Don dep cache va cac dang ky cu...
rmdir /s /q "%LocalAppData%\Microsoft\Office\16.0\Wef" >nul 2>&1
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddin_GitHub" /f >nul 2>&1
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "HoSoAddinDev" /f >nul 2>&1
REG DELETE "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /f >nul 2>&1

echo [4/4] Dang dang ky Trusted Catalog (Local Path)...
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Id" /t REG_SZ /d "%CATALOG_GUID%" /f >nul
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Url" /t REG_SZ /d "%INSTALL_DIR%" /f >nul
REG ADD "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%CATALOG_GUID%" /v "Flags" /t REG_DWORD /d 1 /f >nul

echo.
echo ==================================================
echo DA CAI DAT XONG! 
echo.
echo DUONG DAN THU MUC CUA BAN LA: 
echo %INSTALL_DIR%
echo.
echo NEU TRONG WORD VAN CHUA THAY ADD-IN, HAY LAM THEO BUOC NAY:
echo 1. Trong Word: File -^> Options -^> Trust Center -^> Trust Center Settings.
echo 2. Chon "Trusted Add-in Catalogs".
echo 3. Nhap duong dan phia tren vao o "Catalog Url" roi bam "Add Catalog".
echo 4. Tich vao o "Show in Menu" cua dong vua hien ra.
echo 5. Bam OK, tat Word di mo lai la se thay 100%%.
echo ==================================================
pause
