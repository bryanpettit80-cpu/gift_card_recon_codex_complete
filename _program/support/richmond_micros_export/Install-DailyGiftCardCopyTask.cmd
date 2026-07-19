@echo off
setlocal

rem Installs a daily scheduled task on RESSERVER to publish the current gift-card export.
rem Run this on RESSERVER from the synced Dropbox setup folder.
rem The scheduled task runs a local protected copy, not the mutable Dropbox file.

set "TASK_NAME=Gift Card Export Copy to Dropbox"
set "SOURCE_SCRIPT=%~dp0Copy-GiftCardExportToDropbox.cmd"
set "INSTALL_DIR=%ProgramData%\GiftCardRecon\RichmondMicrosExport"
set "SCRIPT_PATH=%INSTALL_DIR%\Copy-GiftCardExportToDropbox.cmd"

if not exist "%SOURCE_SCRIPT%" (
  echo Cannot find "%SOURCE_SCRIPT%".
  echo Make sure this setup folder has synced to the Richmond server Dropbox first.
  exit /b 20
)

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
if errorlevel 1 (
  echo Cannot create local install folder "%INSTALL_DIR%".
  exit /b 21
)

copy /Y "%SOURCE_SCRIPT%" "%SCRIPT_PATH%" >nul
if errorlevel 1 (
  echo Cannot install local task script to "%SCRIPT_PATH%".
  exit /b 22
)

rem Restrict the local executable copy to administrators and SYSTEM.
icacls "%INSTALL_DIR%" /inheritance:r /grant:r "Administrators:(OI)(CI)F" "SYSTEM:(OI)(CI)F" >nul
if errorlevel 1 (
  echo Cannot secure local install folder "%INSTALL_DIR%".
  exit /b 23
)

schtasks /Create /TN "%TASK_NAME%" /TR "\"%SCRIPT_PATH%\"" /SC DAILY /ST 06:35 /F
set "TASK_EXIT=%ERRORLEVEL%"

if not "%TASK_EXIT%"=="0" (
  echo Scheduled task install failed. Exit code: %TASK_EXIT%
  exit /b %TASK_EXIT%
)

echo Installed scheduled task "%TASK_NAME%" for 06:35 daily.
exit /b 0
