@echo off
:: Run-SyncReceipts.bat
:: Double-click this file to sync receipts into the Excel workbook.
:: Place this file in the same folder as Sync-Receipts.ps1

call "%~dp0config.bat"

:: Add -KillExcel below if Excel crashed and is holding the file locked.
PowerShell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "%~dp0Sync-Receipts.ps1" -ReceiptsRoot "%RECEIPTS_ROOT%"

pause
