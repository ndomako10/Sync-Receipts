# Sync-Receipts

[![Tests](https://github.com/ndomako10/Sync-Receipts/actions/workflows/tests.yml/badge.svg)](https://github.com/ndomako10/Sync-Receipts/actions/workflows/tests.yml)
[![GitHub release](https://img.shields.io/github/v/release/ndomako10/Sync-Receipts)](https://github.com/ndomako10/Sync-Receipts/releases/latest)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](LICENSE)
[![PowerShell 5.0+](https://img.shields.io/badge/PowerShell-5.0%2B-blue)](https://github.com/PowerShell/PowerShell)

A PowerShell automation tool that syncs receipt files into a formatted Excel workbook. Metadata is encoded directly in receipt filenames -- no manual data entry.

## How It Works

Receipt files are named with embedded metadata:

```
YYMMDD Vendor $Amount Method [Account].ext
```

The `Account` field is the last 4 digits of the card or account number. Two special tokens are also accepted:

| Token | Meaning |
|-------|---------|
| `xxxx` | Last 4 is known but intentionally omitted for privacy |
| `----` | Last 4 is unknown (e.g. receipt did not print it) |

Cash receipts carry no account number.

Examples:
```
260316 Sunoco $5.27 Card 9080.pdf
260313 Walmart -$58.22 Card 3232.pdf
260301 Landlord -$1200.00 Checking 4455.pdf
260310 CVS -$12.00 Cash.pdf
260315 Amazon -$34.99 Card xxxx.pdf
260320 Costco -$67.50 Card ----.pdf
```

Running a `.bat` launcher triggers the PowerShell script, which parses every receipt in the target month folder and writes a formatted table into a per-year workbook (e.g. `2026.xlsx`) -- one sheet per month (e.g. `2603` for March 2026).

## Prerequisites

- Windows 10/11
- PowerShell 5.0+
- Microsoft Excel 2016+

## Setup

1. Copy `config.template.bat` to `config.bat`
2. Edit `config.bat` and set:
   - `RECEIPTS_ROOT` -- path to the folder containing your year/month receipt subfolders
3. Copy `Accounts.template.xlsx` to `RECEIPTS_ROOT\Accounts.xlsx`, replace the example rows with your own accounts
4. Optionally edit `Categories.json` in the script folder to customise your categories
5. Run the script -- per-year workbooks (e.g. `2026.xlsx`) are created automatically in `RECEIPTS_ROOT`

## Folder Structure

The script files and the data are kept in separate locations. `RECEIPTS_ROOT` is set in `config.bat`.

```
Script files (e.g. C:\Scripts\Sync-Receipts\):
    config.bat               <- gitignored; sets RECEIPTS_ROOT
    config.template.bat
    Accounts.template.xlsx   <- copy to RECEIPTS_ROOT\Accounts.xlsx and fill in your accounts
    Categories.json          <- category/subcategory definitions; edit to customise
    Sync-Receipts.ps1
    Run-SyncReceipts.bat
    Run-SyncAllReceipts.bat
    Kill-Excel.bat           <- force-closes hung EXCEL.EXE processes

RECEIPTS_ROOT (e.g. \\Server\Share\Receipts\):
    2026.xlsx                <- created automatically on first sync for that year
    2025.xlsx
    Accounts.xlsx            <- account lookup table (Last 4, Holder, Institution, Network, Type)
    2026\
        2603 - March\
            260301 Vendor $10.00 Card 1234.pdf
            ...
        2602 - February\
            ...
    2025\
        ...
```

## Usage

| File | Action |
|------|--------|
| `Run-SyncReceipts.bat` | Sync the current month (reads `RECEIPTS_ROOT` from `config.bat`) |
| `Run-SyncAllReceipts.bat` | Sync all month folders across all years (reads `RECEIPTS_ROOT` from `config.bat`) |

To force-close a crashed Excel instance holding the file locked, double-click `Kill-Excel.bat`, run the **Sync: Current month (KillExcel)** VS Code task, or add `-KillExcel` to the PowerShell command inside `Run-SyncReceipts.bat`.

## Script Parameters

| Parameter | Description |
|-----------|-------------|
| `-ReceiptsRoot` | **Required.** Path to the folder containing year subfolders and per-year workbooks. Set via `config.bat`. |
| `-YearMonth` | YYMM to sync (e.g. `2603`). Defaults to current month. |
| `-Year` | 4-digit year (e.g. `2026`). Syncs all month folders under that year. Mutually exclusive with `-YearMonth` and `-All`. |
| `-WorkbookPath` | Full path to a specific `.xlsx` to write into. Overrides the default per-year path (e.g. `2026.xlsx`). Useful for testing. |
| `-All` | Sync every month folder under every year folder. |
| `-KillExcel` | Kill any running `EXCEL.EXE` before starting. |

## Workbook Structure

### Account Sheet
No longer used by the script. Account data is now read from `Accounts.xlsx` in `RECEIPTS_ROOT`. The sheet in the year workbook can be kept for reference or removed.

### Accounts.xlsx
A dedicated Excel workbook at `RECEIPTS_ROOT\Accounts.xlsx`. Copy from `Accounts.template.xlsx` and fill in your accounts. The script reads **Last 4** (column A) for validation; all other columns are for human reference only.

| Column | Header | Notes |
|--------|--------|-------|
| A | Last 4 | 4-digit account identifier used by the script |
| B | Holder | Account owner |
| C | Institution | Bank, credit union, or fintech (e.g. MFCU, Chase, Apple Pay) |
| D | Network | Visa, Mastercard, Amex, Discover -- **blank for bank accounts and EBT** |
| E | Type | `Credit`, `Debit`, `Checking`, `Savings`, or `EBT` |

**Notes:**
- Bank accounts (Checking/Savings ACH): leave Network blank; set Type to `Checking` or `Savings`
- The same Last 4 may appear on multiple rows (e.g. a bank account number shared across Checking and Savings)
- Cash: no entry needed in `Accounts.xlsx` -- Cash receipts carry no account number in the filename
- EBT: leave Network blank, set Type to `EBT`

If `Accounts.xlsx` is absent, the script falls back to the Account sheet in the year workbook with a deprecation warning. If neither exists, account validation is skipped.

### Category Sheet
No longer used by the script. Category and subcategory data is now read from `Categories.json` in the script folder. The sheet can be kept for reference or removed.

### Month Sheets (e.g. `2603`)
Created or overwritten on each run. Contains a 9-column table:

| Column | Header | Notes |
|--------|--------|-------|
| A | File Name | Hyperlinked to the receipt file |
| B | Date | Formatted `d-mmm` |
| C | Vendor | |
| D | Amount | Currency formatted; negative = expense |
| E | Method | `Card`, `Cash`, `Checking`, or `Savings` |
| F | Account | 4-digit number, or `xxxx` (obfuscated) / `----` (unknown), text formatted |
| G | Category | Dropdown from Category sheet |
| H | Subcategory | Dropdown filtered by selected Category via `INDIRECT` |
| I | Flag | Parse errors; account flags: `Account obfuscated`, `Account unknown`, `Account not in Accounts.xlsx` |

## License

GNU General Public License v3.0 -- see [LICENSE](LICENSE).
