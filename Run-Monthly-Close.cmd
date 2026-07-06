@echo off
setlocal
cd /d "%~dp0"
if not exist "_program\.venv\Scripts\python.exe" goto setup
"_program\.venv\Scripts\python.exe" --version >nul 2>&1
if errorlevel 1 goto setup
goto run

:setup
where pwsh >nul 2>&1
if errorlevel 1 (
  powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\_program\install.ps1"
) else (
  pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\_program\install.ps1"
)
if errorlevel 1 (
  echo.
  echo Setup failed. See the message above.
  pause
  exit /b 1
)

:run
where pwsh >nul 2>&1
if errorlevel 1 (
  powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\_program\run_monthly_close.ps1" %*
) else (
  pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\_program\run_monthly_close.ps1" %*
)
set "exitcode=%errorlevel%"
echo.
echo Open the output folder to find the finished workbook:
echo %cd%\Output
echo.
pause
exit /b %exitcode%
