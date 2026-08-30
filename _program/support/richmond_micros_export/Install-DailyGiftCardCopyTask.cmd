@echo off
setlocal

rem This synced Dropbox file is deliberately not an installer. A Dropbox writer
rem must never receive authority to alter a scheduled task or run a verifier.
echo SECURITY ERROR: Do not install the Richmond gift-card task from the synced Dropbox folder.
echo Run the protected local Trusted-InstallDailyGiftCardCopyTask.ps1 from the approved local release directory.
echo See the Richmond Micros support README for the required provisioning and command.
exit /b 26
