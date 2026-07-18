@echo off
setlocal

rem Installs a daily scheduled task on RESSERVER to publish the current gift-card export.
rem Run this on RESSERVER from the synced Dropbox setup folder.

set "TASK_NAME=Gift Card Export Copy to Dropbox"
set "SOURCE_SCRIPT=%~dp0Copy-GiftCardExportToDropbox.cmd"
set "SAFE_ACTION=%SystemRoot%\System32\cmd.exe /d /c exit 0"

rem Fail closed before any Dropbox or staging operation can fail. Replacing the
rem same-named task with a harmless local action neutralizes legacy installers
rem that executed the mutable synced copy directly.
call schtasks /Create /TN "%TASK_NAME%" /TR "%SAFE_ACTION%" /SC DAILY /ST 06:35 /F
set "TASK_EXIT=%ERRORLEVEL%"

if not "%TASK_EXIT%"=="0" (
  echo SECURITY ERROR: Cannot neutralize the existing scheduled task. Exit code: %TASK_EXIT%
  echo The task may still execute an unsafe Dropbox action. Correct it manually before retrying.
  exit /b 25
)

if not defined LOCALAPPDATA (
  echo Cannot install the task because LOCALAPPDATA is not defined.
  exit /b 21
)

set "INSTALL_DIR=%LOCALAPPDATA%\GiftCardRecon\RichmondMicrosExport"
set "SCRIPT_PATH=%INSTALL_DIR%\Copy-GiftCardExportToDropbox.cmd"

if not exist "%SOURCE_SCRIPT%" (
  echo Cannot find "%SOURCE_SCRIPT%".
  echo Make sure this setup folder has synced to the Richmond server Dropbox first.
  exit /b 20
)

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
if errorlevel 1 (
  echo Cannot create the private task-script folder "%INSTALL_DIR%".
  exit /b 22
)

copy /B /Y "%SOURCE_SCRIPT%" "%SCRIPT_PATH%" >nul
if errorlevel 1 (
  echo Cannot install the task script at "%SCRIPT_PATH%".
  exit /b 23
)

set "GIFT_CARD_SOURCE_SCRIPT=%SOURCE_SCRIPT%"
set "GIFT_CARD_INSTALLED_SCRIPT=%SCRIPT_PATH%"
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "$sha256 = [System.Security.Cryptography.SHA256]::Create(); try { $sourceHash = [System.BitConverter]::ToString($sha256.ComputeHash([System.IO.File]::ReadAllBytes($env:GIFT_CARD_SOURCE_SCRIPT))); $installedHash = [System.BitConverter]::ToString($sha256.ComputeHash([System.IO.File]::ReadAllBytes($env:GIFT_CARD_INSTALLED_SCRIPT))); if ($sourceHash -ne $installedHash) { Write-Error 'Installed task script failed SHA-256 verification.'; exit 1 } } finally { $sha256.Dispose() }"
if errorlevel 1 (
  del /Q "%SCRIPT_PATH%" >nul 2>&1
  echo Installed task script failed SHA-256 verification.
  exit /b 24
)

call schtasks /Create /TN "%TASK_NAME%" /TR "\"%SCRIPT_PATH%\"" /SC DAILY /ST 06:35 /F
set "TASK_EXIT=%ERRORLEVEL%"

if not "%TASK_EXIT%"=="0" (
  echo Secure scheduled task activation failed. Exit code: %TASK_EXIT%
  echo The installer did not restore the mutable Dropbox action.
  exit /b %TASK_EXIT%
)

echo Installed scheduled task "%TASK_NAME%" for 06:35 daily.
echo Verified private task script: "%SCRIPT_PATH%"
exit /b 0
