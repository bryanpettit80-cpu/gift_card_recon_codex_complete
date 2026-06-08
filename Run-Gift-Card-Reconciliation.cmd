@echo off
setlocal
cd /d "%~dp0"
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File ".\Run-Weekly-Reconciliation.ps1"
echo.
pause
