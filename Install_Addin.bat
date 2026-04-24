@echo off
setlocal
chcp 65001 >nul

echo ==================================================
echo CHUONG TRINH CAI DAT ADD-IN (FINAL FIX)
echo ==================================================

:: 1. Tao file script PowerShell tam thoi de xu ly chinh xac
set "PS_SCRIPT=%TEMP%\install_hoso.ps1"

echo $appdata = [System.Environment]::GetFolderPath('ApplicationData') > "%PS_SCRIPT%"
echo $installDir = Join-Path $appdata "HoSoAddin" >> "%PS_SCRIPT%"
echo if (!(Test-Path $installDir)) { New-Item -ItemType Directory -Path $installDir -Force } >> "%PS_SCRIPT%"
echo. >> "%PS_SCRIPT%"
echo # Tai manifest tu GitHub >> "%PS_SCRIPT%"
echo $url = "https://raw.githubusercontent.com/buiquangtrung2012-ops/HoSoAddin/main/manifest.xml" >> "%PS_SCRIPT%"
echo $dest = Join-Path $installDir "manifest.xml" >> "%PS_SCRIPT%"
echo (New-Object Net.WebClient).DownloadFile($url, $dest) >> "%PS_SCRIPT%"
echo. >> "%PS_SCRIPT%"
echo # Dang ky Registry >> "%PS_SCRIPT%"
echo $guid = "{B1A9908E-1C4F-40E2-9EED-7C919D12DF01}" >> "%PS_SCRIPT%"
echo $wefPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$guid" >> "%PS_SCRIPT%"
echo if (!(Test-Path $wefPath)) { New-Item -Path $wefPath -Force } >> "%PS_SCRIPT%"
echo Set-ItemProperty -Path $wefPath -Name "Id" -Value $guid >> "%PS_SCRIPT%"
echo Set-ItemProperty -Path $wefPath -Name "Url" -Value $installDir >> "%PS_SCRIPT%"
echo Set-ItemProperty -Path $wefPath -Name "Flags" -Value 1 -Type DWord >> "%PS_SCRIPT%"
echo. >> "%PS_SCRIPT%"
echo # Word Trusted Location >> "%PS_SCRIPT%"
echo $wordLoc = "HKCU:\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\HoSoAddin" >> "%PS_SCRIPT%"
echo if (!(Test-Path $wordLoc)) { New-Item -Path $wordLoc -Force } >> "%PS_SCRIPT%"
echo Set-ItemProperty -Path $wordLoc -Name "Path" -Value $installDir >> "%PS_SCRIPT%"
echo Set-ItemProperty -Path $wordLoc -Name "AllowSubfolders" -Value 1 -Type DWord >> "%PS_SCRIPT%"
echo Set-ItemProperty -Path $wordLoc -Name "Description" -Value "HoSoAddin Catalog" >> "%PS_SCRIPT%"
echo. >> "%PS_SCRIPT%"
echo # Xoa cache Office >> "%PS_SCRIPT%"
echo $cachePath = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef" >> "%PS_SCRIPT%"
echo if (Test-Path $cachePath) { Remove-Item -Recurse -Force $cachePath } >> "%PS_SCRIPT%"

:: 2. Chay script vua tao
powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%"

:: 3. Xoa file tam
del "%PS_SCRIPT%"

echo.
echo ==================================================
echo DA CAI DAT XONG! 
echo.
echo HAY LAM THEO BUOC NAY:
echo 1. Tat het Word va mo lai.
echo 2. Vao tab Insert -^> My Add-ins.
echo 3. Bam nut "Refresh" (lam moi) o goc tren ben phai.
echo 4. Tab "SHARED FOLDER" se hien ra ngay lap tuc.
echo ==================================================
pause
