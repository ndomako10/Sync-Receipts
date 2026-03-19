@echo off
:: Run-SyncMonthReceipts.bat
:: Double-click this file to sync all receipts for a specific month.
:: pushd maps UNC paths to a temporary drive letter so CMD does not error.
pushd "%~dp0"

set "DATE_FORMAT=yyMMdd"
for /f "usebackq tokens=1,* delims==" %%A in (`findstr /v "^#" "%~dp0..\Config\Config.env"`) do if not "%%A"=="" set "%%A=%%B"
if "%WORKBOOKS_ROOT%"=="" set "WORKBOOKS_ROOT=%RECEIPTS_ROOT%"

for /f %%M in ('PowerShell -NoProfile -Command "(Get-Date).ToString('yyMM')"') do set "DEFAULT_MONTH=%%M"
set /p SYNC_MONTH=Enter YYMM to sync [%DEFAULT_MONTH%]:
if "%SYNC_MONTH%"=="" set "SYNC_MONTH=%DEFAULT_MONTH%"

PowerShell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "%~dp0..\Scripts\Sync-Receipts.ps1" -ReceiptsRoot "%RECEIPTS_ROOT%" -WorkbooksRoot "%WORKBOOKS_ROOT%" -DateFormat "%DATE_FORMAT%" -YearMonth %SYNC_MONTH%

pause
popd
