@echo off
setlocal

rem Writes Richmond gift-card export setup status into the synced current folder.
rem Run this on RESSERVER from the synced Dropbox setup folder.

set "TASK_NAME=Gift Card Export Copy to Dropbox"
set "SOURCE=C:\GetLinkedData"
set "DEST=C:\Users\customer\Dropbox\micros_data\RC-Richmond-current"
set "STATUS=%DEST%\GiftCardSetupStatus.txt"

if not exist "%DEST%" mkdir "%DEST%"

(
  echo Richmond Gift Card Export Setup Status
  echo Checked: %DATE% %TIME%
  echo Computer: %COMPUTERNAME%
  echo.
  echo Source folder:
  if exist "%SOURCE%" (
    echo OK - "%SOURCE%"
  ) else (
    echo MISSING - "%SOURCE%"
  )
  echo.
  echo Source files:
  if exist "%SOURCE%\DLYSYSTT.TXT" (
    echo OK - DLYSYSTT.TXT
  ) else (
    echo MISSING - DLYSYSTT.TXT
  )
  if exist "%SOURCE%\TENDER_DETAIL.TXT" (
    echo OK - TENDER_DETAIL.TXT
  ) else (
    echo MISSING - TENDER_DETAIL.TXT
  )
  if exist "%SOURCE%\Micros3700.7z" (
    echo OK - Micros3700.7z
  ) else (
    echo MISSING - Micros3700.7z
  )
  echo.
  echo Destination folder:
  if exist "%DEST%" (
    echo OK - "%DEST%"
  ) else (
    echo MISSING - "%DEST%"
  )
  echo.
  echo Destination files:
  if exist "%DEST%\DLYSYSTT.TXT" (
    echo OK - DLYSYSTT.TXT
  ) else (
    echo MISSING - DLYSYSTT.TXT
  )
  if exist "%DEST%\TENDER_DETAIL.TXT" (
    echo OK - TENDER_DETAIL.TXT
  ) else (
    echo MISSING - TENDER_DETAIL.TXT
  )
  if exist "%DEST%\Micros3700.7z" (
    echo OK - Micros3700.7z
  ) else (
    echo MISSING - Micros3700.7z
  )
  echo.
  echo Scheduled task:
) > "%STATUS%"

schtasks /Query /TN "%TASK_NAME%" /FO LIST /V >> "%STATUS%" 2>&1
set "TASK_EXIT=%ERRORLEVEL%"

(
  echo.
  if "%TASK_EXIT%"=="0" (
    echo TASK_STATUS=INSTALLED
  ) else (
    echo TASK_STATUS=NOT_FOUND_OR_NOT_READABLE
  )
) >> "%STATUS%"

type "%STATUS%"
exit /b 0
