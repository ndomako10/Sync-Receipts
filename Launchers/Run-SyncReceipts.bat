@echo off
:: Run-SyncReceipts.bat
:: Double-click this file to sync receipts into the Excel workbook.
:: pushd maps UNC paths to a temporary drive letter so CMD does not error.
pushd "%~dp0"

set "DATE_FORMAT=yyMMdd"
call "%~dp0..\Config\Config.bat"
if "%WORKBOOKS_ROOT%"=="" set "WORKBOOKS_ROOT=%RECEIPTS_ROOT%"

:: Add -KillExcel below if Excel crashed and is holding the file locked.
PowerShell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "%~dp0..\Scripts\Sync-Receipts.ps1" -ReceiptsRoot "%RECEIPTS_ROOT%" -WorkbooksRoot "%WORKBOOKS_ROOT%" -DateFormat "%DATE_FORMAT%"

pause
popd
