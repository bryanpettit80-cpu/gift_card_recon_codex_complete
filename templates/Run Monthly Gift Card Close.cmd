@echo off
setlocal
title Monthly Gift Card Close
set "OPERATIONS_ROOT=%~dp0."
set "PROGRAM_ROOT=%~dp0Gift Card Reconciliation Automation"
set "RUNNER=%PROGRAM_ROOT%\_program\run_monthly_close.ps1"

echo.
echo MONTHLY GIFT CARD CLOSE
echo =======================
echo The program will scan the Darden inbox and process each valid store-period independently.
echo.

if not exist "%RUNNER%" (
  echo ATTENTION NEEDED: The automation program was not found:
  echo %RUNNER%
  echo Run the operator-asset installer from the program folder or ask for technical help.
  echo.
  pause
  exit /b 2
)

where pwsh >nul 2>&1
if errorlevel 1 (
  powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%RUNNER%" -OperationsRoot "%OPERATIONS_ROOT%" %*
) else (
  pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%RUNNER%" -OperationsRoot "%OPERATIONS_ROOT%" %*
)
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" (
  echo Finished. Review the close statuses and report paths shown above.
) else (
  echo ATTENTION NEEDED: The monthly close did not finish normally.
  echo Review the exact message and diagnostic paths shown above.
)
echo.
echo You may close this window after reviewing the result.
pause
exit /b %EXITCODE%
