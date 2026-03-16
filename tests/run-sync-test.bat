@echo off
call "%~dp0..\config.bat"
powershell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "%~dp0..\Sync-Receipts.ps1" -ReceiptsRoot "%RECEIPTS_ROOT_LOCAL%" -WorkbookPath "%~dp0Receipts (test).xlsx" > "%~dp0sync-output.txt" 2>&1
