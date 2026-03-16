@echo off
:: Run-SyncAllReceipts.bat
:: Double-click this file to sync all receipt months into the Excel workbook.
:: Place this file in the same folder as Sync-Receipts.ps1
:: Note: -ReceiptsRoot is omitted because the script defaults to its own folder,
::       which is the Receipts root when all files are placed there together.

PowerShell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "%~dp0Sync-Receipts.ps1" -All

pause
