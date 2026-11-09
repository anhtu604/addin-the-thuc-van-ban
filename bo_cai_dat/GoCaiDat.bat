@echo off
chcp 65001 >nul
cls
echo GO BO ADD-IN CHUAN HOA THE THUC
echo.
:choword
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  echo [!] Hay LUU va DONG het cua so Word, roi bam phim bat ky...
  pause >nul
  goto choword
)
set "T=%APPDATA%\Microsoft\Word\STARTUP\ChuanHoa_VanBan_Rieng.dotm"
if exist "%T%" del /F /Q "%T%" >nul 2>&1
echo Da go add-in khoi Word.
pause
