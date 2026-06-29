@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" goto setup
".venv\Scripts\python.exe" --version >nul 2>&1
if errorlevel 1 goto setup
goto run

:setup
where pwsh >nul 2>&1
if errorlevel 1 (
  powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\install.ps1"
) else (
  pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\install.ps1"
)
if errorlevel 1 (
  echo.
  echo Setup failed. See the message above.
  pause
  exit /b 1
)

:run
if not exist ".\output" mkdir ".\output"
".venv\Scripts\python.exe" -m gift_card_recon.auto_run --input-root ".\input" --output-dir ".\output"
set "exitcode=%errorlevel%"
echo.
echo Open the output folder to find the finished workbook(s):
echo %cd%\output
echo.
pause
exit /b %exitcode%
