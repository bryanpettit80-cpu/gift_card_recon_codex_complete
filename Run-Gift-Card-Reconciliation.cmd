@echo off
setlocal
cd /d "%~dp0"
where pwsh >nul 2>&1
if errorlevel 1 (
  powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\_program\run_weekly.ps1" %*
) else (
  pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\_program\run_weekly.ps1" %*
)
set "exitcode=%errorlevel%"
echo.
if "%exitcode%"=="0" (
  echo Weekly reconciliation finished successfully. Review the results and file paths shown above.
) else (
  echo Weekly reconciliation did not complete. Review the messages above.
)
echo.
pause
exit /b %exitcode%
