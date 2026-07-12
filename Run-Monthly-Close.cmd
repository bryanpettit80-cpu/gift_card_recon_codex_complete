@echo off
setlocal
cd /d "%~dp0"
where pwsh >nul 2>&1
if errorlevel 1 (
  powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\_program\run_monthly_close.ps1" %*
) else (
  pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\_program\run_monthly_close.ps1" %*
)
set "exitcode=%errorlevel%"
echo.
if "%exitcode%"=="0" (
  echo Monthly close command finished successfully. Review the status and file paths shown above.
) else (
  echo Monthly close did not complete. Review the messages and diagnostic paths shown above.
)
echo.
pause
exit /b %exitcode%
