@echo off
:: Setup.bat
:: Run this once on a new machine to configure Sync-Receipts.
:: Creates Config.bat, copies Accounts.xlsx, and creates shortcuts in RECEIPTS_ROOT.

:: pushd maps UNC paths to a temporary drive letter so CMD does not error.
pushd "%~dp0"

PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Initialize-SyncReceipts.ps1"

pause
popd
