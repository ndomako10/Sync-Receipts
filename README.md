# Sync-Receipts

A PowerShell automation tool that syncs receipt files into a formatted Excel workbook. Metadata is encoded directly in receipt filenames — no manual data entry.

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

Running a `.bat` launcher triggers the PowerShell script, which parses every receipt in the target month folder and writes a formatted table into `Receipts.xlsx` — one sheet per month (e.g. `2603` for March 2026).

## Prerequisites

- Windows 10/11
- PowerShell 5.0+
- Microsoft Excel 2016+

## Setup

1. Copy `config.template.bat` to `config.bat`
2. Edit `config.bat` and set:
   - `RECEIPTS_ROOT` — path to the folder containing `Receipts.xlsx` and your year/month subfolders
   - `DOWNLOADS_DIR` — your Downloads folder (used by `Deploy-SyncReceipts.bat`)
3. Make sure `Receipts.xlsx` exists in `RECEIPTS_ROOT` with an **Account** sheet and a **Category** sheet (see [Workbook Structure](#workbook-structure))

## Folder Structure

```
RECEIPTS_ROOT\
    Receipts.xlsx
    Sync-Receipts.ps1
    Run-SyncReceipts.bat
    Run-SyncAllReceipts.bat
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
| `Run-SyncReceipts.bat` | Sync the current month |
| `Run-SyncAllReceipts.bat` | Sync all month folders across all years |

Add `-KillExcel` to the PowerShell command inside `Run-SyncReceipts.bat` if Excel crashed and is holding the file locked.

## Script Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-ReceiptsRoot` | Script's own folder | Path to the root Receipts folder |
| `-YearMonth` | Current month (yyMM) | Specific month to sync, e.g. `2603` |
| `-All` | — | Sync every month folder under every year folder |
| `-KillExcel` | — | Kill any running `EXCEL.EXE` before starting |

## Workbook Structure

### Account Sheet
Column A, row 2 downward: 4-digit account numbers. Used to validate accounts parsed from filenames. If the sheet is missing, validation is skipped.

### Category Sheet
Row 1: category headers (e.g. `Food`, `Housing`, `Gas`). Rows 2+: subcategories under each column. Used to populate Category dropdowns and named ranges. If the sheet is missing, dropdowns are skipped.

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
| H | Subcategory | Dropdown (see known issues) |
| I | Flag | Parse errors, unknown accounts |

## Known Issues

See [ISSUES.md](ISSUES.md).

## License

MIT — see [LICENSE](LICENSE).
