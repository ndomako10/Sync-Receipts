@echo off
:: Config.template.bat - copy this to Config.bat and fill in your values

:: Path to the root Receipts folder (network share or local path)
set "RECEIPTS_ROOT=C:\Users\YOUR-USERNAME\Documents\Receipts"

:: Local drive equivalent of RECEIPTS_ROOT, used by VS Code tasks
:: to avoid UNC path latency. Set to the same folder via a local drive letter.
set "RECEIPTS_ROOT_LOCAL=C:\Users\YOUR-USERNAME\Documents\Receipts"

:: Receipt filename date format -- .NET ParseExact format string (default: yyMMdd)
:: yyMMdd = 260316, yyyyMMdd = 20260316, yy-MM-dd = 26-03-16, MMddyy = 031626
set "DATE_FORMAT=yyMMdd"
