@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\install.ps1"
)
if not exist ".\output" mkdir ".\output"
".venv\Scripts\python.exe" -m gift_card_recon.auto_run --input-root ".\input" --output-dir ".\output"
echo.
echo Open the output folder to find the finished workbook(s):
echo %cd%\output
echo.
pause
