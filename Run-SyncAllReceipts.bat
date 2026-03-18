@echo off
:: Run-SyncAllReceipts.bat
:: Double-click this file to sync all receipt months into the Excel workbook.

:: pushd maps UNC paths to a temporary drive letter so CMD does not error.
pushd "%~dp0"

call "%~dp0config.bat"

PowerShell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "%~dp0Sync-Receipts.ps1" -ReceiptsRoot "%RECEIPTS_ROOT%" -All

pause
popd
