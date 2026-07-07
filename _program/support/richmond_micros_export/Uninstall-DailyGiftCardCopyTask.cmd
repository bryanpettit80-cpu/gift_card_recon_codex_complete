@echo off
setlocal

rem Removes the scheduled copy task created by Install-DailyGiftCardCopyTask.cmd.

set "TASK_NAME=Gift Card Export Copy to Dropbox"
schtasks /Delete /TN "%TASK_NAME%" /F
exit /b %ERRORLEVEL%
