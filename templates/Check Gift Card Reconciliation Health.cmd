@echo off
setlocal
title Gift Card Reconciliation Health Check
set "OPERATIONS_ROOT=%~dp0."
set "PROGRAM_ROOT=%~dp0Gift Card Reconciliation Automation"
set "CHECKER=%PROGRAM_ROOT%\_program\check_operator_health.ps1"

echo.
if not exist "%CHECKER%" (
  echo ATTENTION NEEDED: The operator health checker was not found:
  echo %CHECKER%
  echo Deploy the current program from the clean local Git checkout.
  echo.
  pause
  exit /b 2
)

where pwsh >nul 2>&1
if errorlevel 1 (
  powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%CHECKER%" -OperationsRoot "%OPERATIONS_ROOT%" %*
) else (
  pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%CHECKER%" -OperationsRoot "%OPERATIONS_ROOT%" %*
)
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" (
  echo The Gift Card Reconciliation operator environment is ready.
) else (
  echo ATTENTION NEEDED: One or more health controls are blocked.
)
echo.
pause
exit /b %EXITCODE%
