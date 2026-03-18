@echo off
:: Run-SyncReceipts.bat
:: Double-click this file to sync receipts into the Excel workbook.
:: pushd maps UNC paths to a temporary drive letter so CMD does not error.
pushd "%~dp0"

call "%~dp0..\Config.bat"

:: Add -KillExcel below if Excel crashed and is holding the file locked.
PowerShell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "%~dp0..\Scripts\Sync-Receipts.ps1" -ReceiptsRoot "%RECEIPTS_ROOT%"

pause
popd
