@echo off
setlocal

rem Copies the current Richmond Micros gift-card export files into Dropbox.
rem Run this on RESSERVER after the normal C:\GetLinkedData export completes.

if not defined GIFT_CARD_EXPORT_SOURCE set "GIFT_CARD_EXPORT_SOURCE=C:\GetLinkedData"
if not defined GIFT_CARD_EXPORT_DEST set "GIFT_CARD_EXPORT_DEST=C:\Users\customer\Dropbox\micros_data\RC-Richmond-current"
if not defined GIFT_CARD_EXPORT_LOG set "GIFT_CARD_EXPORT_LOG=%GIFT_CARD_EXPORT_SOURCE%\GiftCardCopy.log"

if not exist "%GIFT_CARD_EXPORT_SOURCE%\DLYSYSTT.TXT" (
  echo Missing required source file: "%GIFT_CARD_EXPORT_SOURCE%\DLYSYSTT.TXT"
  exit /b 10
)

if not exist "%GIFT_CARD_EXPORT_SOURCE%\TENDER_DETAIL.TXT" (
  echo Missing required source file: "%GIFT_CARD_EXPORT_SOURCE%\TENDER_DETAIL.TXT"
  exit /b 11
)

if not exist "%GIFT_CARD_EXPORT_DEST%" mkdir "%GIFT_CARD_EXPORT_DEST%"

robocopy "%GIFT_CARD_EXPORT_SOURCE%" "%GIFT_CARD_EXPORT_DEST%" Micros3700.7z DLYSYSTT.TXT TENDER_DETAIL.TXT /R:2 /W:5 /NP /LOG+:"%GIFT_CARD_EXPORT_LOG%"
set "ROBOCOPY_EXIT=%ERRORLEVEL%"

if %ROBOCOPY_EXIT% GEQ 8 (
  echo Gift card export copy failed. Robocopy exit code: %ROBOCOPY_EXIT%
  exit /b %ROBOCOPY_EXIT%
)

echo Gift card export files copied to "%GIFT_CARD_EXPORT_DEST%".
exit /b 0
