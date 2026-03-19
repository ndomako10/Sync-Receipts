@echo off
:: Run-SyncAllReceipts.bat
:: Double-click this file to sync all receipt months into the Excel workbook.

:: pushd maps UNC paths to a temporary drive letter so CMD does not error.
pushd "%~dp0"

set "DATE_FORMAT=yyMMdd"
for /f "usebackq tokens=1,* delims==" %%A in (`findstr /v "^#" "%~dp0..\Config\Config.env"`) do if not "%%A"=="" set "%%A=%%B"
if "%WORKBOOKS_ROOT%"=="" set "WORKBOOKS_ROOT=%RECEIPTS_ROOT%"

PowerShell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "%~dp0..\Scripts\Sync-Receipts.ps1" -ReceiptsRoot "%RECEIPTS_ROOT%" -WorkbooksRoot "%WORKBOOKS_ROOT%" -DateFormat "%DATE_FORMAT%" -All

pause
popd
