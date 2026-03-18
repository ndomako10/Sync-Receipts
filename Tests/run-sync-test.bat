@echo off
call "%~dp0..\Config.bat"
powershell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "%~dp0..\Scripts\Sync-Receipts.ps1" -ReceiptsRoot "%RECEIPTS_ROOT_LOCAL%" -WorkbookPath "%~dp0Receipts (test).xlsx" > "%~dp0sync-output.txt" 2>&1
