@echo off
setlocal

rem Installs a daily scheduled task on RESSERVER to publish the current gift-card export.
rem Run this on RESSERVER from the synced Dropbox setup folder.

set "TASK_NAME=Gift Card Export Copy to Dropbox"
set "SCRIPT_PATH=C:\Users\customer\Dropbox\micros_data\RC-Richmond\_gift_card_current_export_setup\Copy-GiftCardExportToDropbox.cmd"

if not exist "%SCRIPT_PATH%" (
  echo Cannot find "%SCRIPT_PATH%".
  echo Make sure this setup folder has synced to the Richmond server Dropbox first.
  exit /b 20
)

schtasks /Create /TN "%TASK_NAME%" /TR "\"%SCRIPT_PATH%\"" /SC DAILY /ST 06:35 /F
set "TASK_EXIT=%ERRORLEVEL%"

if not "%TASK_EXIT%"=="0" (
  echo Scheduled task install failed. Exit code: %TASK_EXIT%
  exit /b %TASK_EXIT%
)

echo Installed scheduled task "%TASK_NAME%" for 06:35 daily.
exit /b 0
