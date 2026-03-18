@echo off
:: Run-SyncYearReceipts.bat
:: Double-click this file to sync all receipt months for a specific year.
:: pushd maps UNC paths to a temporary drive letter so CMD does not error.
pushd "%~dp0"

set "DATE_FORMAT=yyMMdd"
call "%~dp0..\Config\Config.bat"

for /f %%Y in ('PowerShell -NoProfile -Command "(Get-Date).Year"') do set "DEFAULT_YEAR=%%Y"
set /p SYNC_YEAR=Enter 4-digit year to sync [%DEFAULT_YEAR%]:
if "%SYNC_YEAR%"=="" set "SYNC_YEAR=%DEFAULT_YEAR%"

PowerShell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "%~dp0..\Scripts\Sync-Receipts.ps1" -ReceiptsRoot "%RECEIPTS_ROOT%" -DateFormat "%DATE_FORMAT%" -Year %SYNC_YEAR%

pause
popd
