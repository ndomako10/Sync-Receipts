# Sync-Receipts

[![Tests](https://github.com/ndomako10/Sync-Receipts/actions/workflows/tests.yml/badge.svg)](https://github.com/ndomako10/Sync-Receipts/actions/workflows/tests.yml)
[![GitHub release](https://img.shields.io/github/v/release/ndomako10/Sync-Receipts)](https://github.com/ndomako10/Sync-Receipts/releases/latest)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](LICENSE)
[![PowerShell 5.0+](https://img.shields.io/badge/PowerShell-5.0%2B-blue)](https://github.com/PowerShell/PowerShell)

A PowerShell automation tool that syncs receipt files into a formatted Excel workbook. Metadata is encoded directly in receipt filenames -- no manual data entry.

## How It Works

Receipt files are named with embedded metadata:

```
yyMMdd Vendor $Amount [Method [Account]].ext
```

`yyMMdd` is a [.NET ParseExact format string](https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings): `yy` = 2-digit year, `MM` = month, `dd` = day.

`Method` and `Account` are optional. If omitted, the row is written to the workbook with both fields blank and flagged `Method missing` so the user can fill them in.

The `Account` field is the last 4 digits of the card or account number. Two special tokens are also accepted:

| Token | Meaning |
|-------|---------|
| `xxxx` | Last 4 is known but intentionally omitted for privacy |
| `----` | Last 4 is unknown (e.g. receipt did not print it) |

Cash receipts carry no account number.

Examples:
```
260301 Landlord $1200.00 Checking 4455.pdf
260305 CVS $12.00 Cash.pdf
260310 Sunoco $5.27 Card 9080.pdf
260315 Amazon $34.99 Card xxxx.pdf
260320 Costco $67.50 Card ----.pdf
260325 Employer -$2500.00 Savings 7890.pdf
260328 Vendor $5.00.pdf
```

Running a `.bat` launcher triggers the PowerShell script, which parses every receipt in the target month folder and writes a formatted table into a per-year workbook (e.g. `2026.xlsx`) -- one sheet per month (e.g. `2603` for March 2026).

## Prerequisites

- Windows 10/11
- PowerShell 5.0+
- Microsoft Excel 2016+

## Setup

1. Double-click `Setup.bat` -- it will:
   - Check that PowerShell 5.0+ and Excel are installed
   - Prompt for `RECEIPTS_ROOT` and `WORKBOOKS_ROOT` (defaults to `RECEIPTS_ROOT`), then create `Config.env` (skipped if `Config.env` already exists)
   - Create the `RECEIPTS_ROOT` and `WORKBOOKS_ROOT` folders if they do not exist (prompts for confirmation)
   - Copy `Config\Accounts.template.xlsx` to `Config\Accounts.xlsx` (skipped if already present)
   - Create `Run Sync Receipts.lnk`, `Run Sync Month Receipts.lnk`, `Run Sync Year Receipts.lnk`, and `Run Sync All Receipts.lnk` shortcuts in `RECEIPTS_ROOT`
2. Open `Config\Accounts.xlsx` and replace the example rows with your own accounts
3. Copy `Config\Categories.template.json` to `Config\Categories.json` and edit to customise your categories (skipped if already present)
4. For the first run, use the `Run Sync All Receipts.lnk` shortcut in `RECEIPTS_ROOT` to sync all existing receipts across all years into workbooks
5. After the initial sync, use `Run Sync Receipts.lnk` for day-to-day syncing of the current month. Use the month and year launchers when correcting past entries.

`Setup.bat` is safe to re-run -- it skips steps that are already complete.

## Folder Structure

The script files and the data are kept in separate locations. `RECEIPTS_ROOT` is set in `Config\Config.env`.

```
Script files (e.g. C:\Scripts\Sync-Receipts\):
    Setup.bat                <- one-time setup launcher
    Config\
        Config.env               <- gitignored; sets RECEIPTS_ROOT
        Config.template.env
        Accounts.template.xlsx  <- default accounts template; committed to git
        Accounts.xlsx           <- gitignored; your personal accounts
        Categories.template.json <- default categories template; committed to git
        Categories.json          <- gitignored; your personal categories
    Scripts\
        Initialize-SyncReceipts.ps1
        Sync-Receipts.ps1
    Launchers\
        Run-SyncReceipts.bat
        Run-SyncAllReceipts.bat
    Kill-Excel.bat           <- force-closes hung EXCEL.EXE processes (not in repo)

RECEIPTS_ROOT (e.g. \\Server\Share\Receipts\):
    2026.xlsx                <- created automatically on first sync for that year
    2025.xlsx
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

`Launchers\Run-SyncReceipts.bat` is the standard day-to-day launcher. The other launchers are for initial setup and correcting past entries.

| File | Action |
|------|--------|
| `Launchers\Run-SyncReceipts.bat` | **Day-to-day.** Sync the current month. |
| `Launchers\Run-SyncMonthReceipts.bat` | Prompt for a YYMM month to sync; defaults to current month. |
| `Launchers\Run-SyncYearReceipts.bat` | Prompt for a 4-digit year to sync; defaults to current year. |
| `Launchers\Run-SyncAllReceipts.bat` | Sync all month folders across all years. |

To force-close a crashed Excel instance holding the file locked, run the **Sync: Current month (KillExcel)** VS Code task, or add `-KillExcel` to the PowerShell command inside `Launchers\Run-SyncReceipts.bat`.

## Script Parameters

| Parameter | Description |
|-----------|-------------|
| `-ReceiptsRoot` | **Required.** Path to the folder containing year subfolders and receipt files. Set via `Config\Config.env`. |
| `-WorkbooksRoot` | Directory where per-year workbooks (e.g. `2026.xlsx`) are written. Defaults to `ReceiptsRoot`. Set `WORKBOOKS_ROOT` in `Config\Config.env` to store workbooks separately from receipts. |
| `-YearMonth` | YYMM to sync (e.g. `2603`). Defaults to current month. |
| `-Year` | 4-digit year (e.g. `2026`). Syncs all month folders under that year. Mutually exclusive with `-YearMonth` and `-All`. |
| `-WorkbookPath` | Full path to a specific `.xlsx` to write into. Overrides the default per-year path derived from `WorkbooksRoot`. Useful for testing. |
| `-DateFormat` | [.NET ParseExact format string](https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings) for the date portion of receipt filenames. Default: `yyMMdd`. Set `DATE_FORMAT` in `Config\Config.env` to change. Single-digit tokens `M` and `d` are supported only in separator-delimited formats (e.g. `M-d-yy`); use `MM`/`dd` for compact formats. |
| `-All` | Sync every month folder under every year folder. |
| `-KillExcel` | Kill any running `EXCEL.EXE` before starting. |

## Workbook Structure

### Accounts.xlsx
A dedicated Excel workbook at `Config\Accounts.xlsx` (gitignored). Copy from `Config\Accounts.template.xlsx` and fill in your accounts. The script reads **Last 4** (column A) for validation; all other columns are for human reference only.

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

If `Accounts.xlsx` is absent, account validation is skipped.

### Category Sheet
Written and maintained automatically by the script on every run; hidden in the tab bar. Category and subcategory data is sourced from `Config\Categories.json` (gitignored). Copy from `Config\Categories.template.json` to get started.

### Month Sheets (e.g. `2603`)
Created or overwritten on each run. Contains a 9-column table:

| Column | Header | Notes |
|--------|--------|-------|
| A | File Name | Hyperlinked to the receipt file |
| B | Date | Formatted `d-mmm` |
| C | Vendor | Name of the merchant or payee |
| D | Amount | Currency formatted; positive = expense, negative = income |
| E | Method | `Card`, `Cash`, `Checking`, `Savings`, or blank if omitted |
| F | Account | 4-digit number, or `xxxx` (obfuscated) / `----` (unknown), text formatted; blank if Method omitted |
| G | Category | Dropdown; list sourced from the hidden Category sheet (auto-populated) |
| H | Subcategory | Dropdown filtered by selected Category via `INDIRECT` |
| I | Flag | Parse errors: `Could not parse filename`, `Month out of range`, `Day out of range`, `Invalid date`; missing field: `Method missing`; account flags: `Account obfuscated`, `Account unknown`, `Account not in Accounts.xlsx` |

## License

GNU General Public License v3.0 -- see [LICENSE](LICENSE).
