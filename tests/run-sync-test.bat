@echo off
powershell -NoProfile -ExecutionPolicy Bypass -NonInteractive -File "f:\Documents\Sync-Receipts\Sync-Receipts.ps1" -ReceiptsRoot "f:\Documents\Receipts" -WorkbookPath "f:\Documents\Sync-Receipts\tests\Receipts (test).xlsx" > "f:\Documents\Sync-Receipts\tests\sync-output.txt" 2>&1
