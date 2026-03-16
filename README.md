# Sync-Receipts

A PowerShell automation tool that syncs receipt files into a formatted Excel workbook. Metadata is encoded directly in receipt filenames -- no manual data entry.

## How It Works

Receipt files are named with embedded metadata:

```
YYMMDD Vendor $Amount Method [Account].ext
```

Examples:
```
260316 Sunoco $5.27 Card 9080.pdf
260313 Walmart -$58.22 Card 3232.pdf
260301 Landlord -$1200.00 Checking 4455.pdf
```

Running a `.bat` launcher triggers the PowerShell script, which parses every receipt in the target month folder and writes a formatted table into `Receipts.xlsx` -- one sheet per month (e.g. `2603` for March 2026).

## Prerequisites

- Windows 10/11
- PowerShell 5.0+
- Microsoft Excel 2016+

## Setup

1. Copy `config.template.bat` to `config.bat`
2. Edit `config.bat` and set:
   - `RECEIPTS_ROOT` -- path to the folder containing `Receipts.xlsx` and your year/month receipt subfolders
3. Make sure `Receipts.xlsx` exists in `RECEIPTS_ROOT` with an **Account** sheet and a **Category** sheet (see [Workbook Structure](#workbook-structure))

## Folder Structure

The script files and the data are kept in separate locations. `RECEIPTS_ROOT` is set in `config.bat`.

```
Script files (e.g. C:\Scripts\Sync-Receipts\):
    config.bat               <- gitignored; sets RECEIPTS_ROOT
    config.template.bat
    Sync-Receipts.ps1
    Run-SyncReceipts.bat
    Run-SyncAllReceipts.bat

RECEIPTS_ROOT (e.g. \\Server\Share\Receipts\):
    Receipts.xlsx
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

To force-close a crashed Excel instance holding the file locked, run the **Sync: Current month (KillExcel)** VS Code task, or add `-KillExcel` to the PowerShell command inside `Run-SyncReceipts.bat`.

## Script Parameters

| Parameter | Description |
|-----------|-------------|
| `-ReceiptsRoot` | **Required.** Path to the folder containing `Receipts.xlsx` and year subfolders. Set via `config.bat`. |
| `-YearMonth` | YYMM to sync (e.g. `2603`). Defaults to current month. |
| `-All` | Sync every month folder under every year folder. |
| `-KillExcel` | Kill any running `EXCEL.EXE` before starting. |

## Workbook Structure

### Account Sheet
Column A, row 2 downward: 4-digit account numbers. Used to validate accounts parsed from filenames. If the sheet is missing, validation is skipped.

### Category Sheet
Row 1: category headers (e.g. `Food`, `Housing`, `Gas`). Rows 2+: subcategories under each column. Used to populate Category and Subcategory dropdowns via named ranges.

### Month Sheets (e.g. `2603`)
Created or overwritten on each run. Contains a 9-column table:

| Column | Header | Notes |
|--------|--------|-------|
| A | File Name | Hyperlinked to the receipt file |
| B | Date | Formatted `d-mmm` |
| C | Vendor | |
| D | Amount | Currency formatted; negative = expense |
| E | Method | `Card`, `Cash`, `Checking`, or `Savings` |
| F | Account | 4-digit number, text formatted |
| G | Category | Dropdown from Category sheet |
| H | Subcategory | Dropdown filtered by selected Category via `INDIRECT` |
| I | Flag | Parse errors, unknown accounts |

## License

MIT -- see [LICENSE](LICENSE).
