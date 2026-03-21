@echo off
for /f "usebackq tokens=1,* delims==" %%A in (`findstr /v "^#" "%~dp0..\Config\Config.ini"`) do if not "%%A"=="" set "%%A=%%B"
powershell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "%~dp0..\Scripts\Sync-Receipts.ps1" -ReceiptsRoot "%RECEIPTS_ROOT%" -WorkbookPath "%~dp0Receipts (test).xlsx" > "%~dp0sync-output.txt" 2>&1
