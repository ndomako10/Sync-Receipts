@echo off
:: Config.template.bat - copy this to Config.bat and fill in your values

:: Path to the root Receipts folder (network share or local path)
set "RECEIPTS_ROOT=C:\Users\YOUR-USERNAME\Documents\Receipts"

:: Local drive equivalent of RECEIPTS_ROOT, used by Tests/run-sync-test.bat
:: to avoid UNC path latency. Set to the same folder via its drive letter.
set "RECEIPTS_ROOT_LOCAL=C:\Users\YOUR-USERNAME\Documents\Receipts"
