@echo off
chcp 65001 >nul
cls
echo ======================================================
echo   CAI DAT ADD-IN CHUAN HOA THE THUC VAN BAN
echo ======================================================
echo.
set "SCRIPT_DIR=%~dp0"
set "SRC=%SCRIPT_DIR%ChuanHoa_VanBan_Rieng.dotm"
set "TGT=%APPDATA%\Microsoft\Word\STARTUP"

if not exist "%SRC%" (
  echo [!] Khong thay file ChuanHoa_VanBan_Rieng.dotm. Hay giai nen toan bo thu muc roi chay lai.
  pause & exit /b 1
)

:choword
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  echo [!] Word dang mo. Hay LUU va DONG het cua so Word, roi bam phim bat ky de tiep tuc...
  pause >nul
  goto choword
)

if not exist "%TGT%" mkdir "%TGT%" 2>nul
if exist "%TGT%\ChuanHoa_VanBanDang.dotm" (
  del /F /Q "%TGT%\ChuanHoa_VanBanDang.dotm" >nul 2>&1
  echo [+] Da go ban goc ChuanHoa_VanBanDang.dotm de tranh trung nut tren Ribbon.
)
copy /Y "%SRC%" "%TGT%\ChuanHoa_VanBan_Rieng.dotm" >nul
if not exist "%TGT%\ChuanHoa_VanBan_Rieng.dotm" (
  echo [!] Khong chep duoc file vao thu muc STARTUP cua Word.
  pause & exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -LiteralPath '%TGT%\ChuanHoa_VanBan_Rieng.dotm'" >nul 2>&1

for %%V in (15.0 16.0) do (
  reg add "HKCU\Software\Microsoft\Office\%%V\Word\Security\Trusted Locations\VPDU_WordStartup" /v "Path" /t REG_SZ /d "%TGT%\" /f >nul 2>&1
)
reg add "HKCU\Software\VB and VBA Program Settings\OfficePartySuite\TheThuc" /v "DonVi" /t REG_SZ /d "AUTO" /f >nul 2>&1

echo [+] Da cai ChuanHoa_VanBan_Rieng.dotm (khoi: Tu dong nhan dien).
echo.
echo Mo lai Word: the Home co nhom [THE THUC VAN BAN], nut [Chuan Hoa The Thuc].
echo Thong tin don vi da nhap truoc day duoc giu nguyen.
echo ======================================================
pause
