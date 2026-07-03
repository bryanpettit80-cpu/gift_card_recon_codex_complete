@echo off
setlocal
cd /d "%~dp0"
if not exist ".\input\gmail_activity" mkdir ".\input\gmail_activity"
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
echo.
echo Put the Gmail-downloaded Gift Card Activity attachments in:
echo %cd%\input\gmail_activity
echo.
".venv\Scripts\python.exe" -m gift_card_recon.gmail_activity_import --source-dir ".\input\gmail_activity" --input-root ".\input"
set "exitcode=%errorlevel%"
echo.
pause
exit /b %exitcode%
