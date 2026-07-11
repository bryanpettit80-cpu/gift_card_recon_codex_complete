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
if "%exitcode%"=="0" (
  echo Monthly close completed. Workbook and PDF reports are in:
  echo %cd%\Output\Monthly Close
) else (
  echo Monthly close did not complete. Review diagnostics, if created, are in:
  echo %cd%\Output\Review Required
)
echo.
pause
exit /b %exitcode%
