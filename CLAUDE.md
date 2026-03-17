# Claude Code Context -- Sync-Receipts

## What this project is

A PowerShell script (`Sync-Receipts.ps1`) that uses Excel COM automation to parse receipt filenames and sync metadata into a formatted Excel workbook. See README.md for full usage details.

## Coding rules

- **Always add error handling and debug output** -- every new block needs `try/catch` and `Write-Log` calls (use `-Tag WARN` for warnings, `-Tag ERROR` for errors, `-Tag VERB` for diagnostic detail)
- **Never use `$variable:` in double-quoted strings** -- PowerShell interprets the colon as a drive separator; use `${variable}:` instead
- **No smart quotes or em-dashes** -- the file must be pure ASCII. Non-ASCII characters break PowerShell parsing on the network share. Verify after any edit: `[System.Text.Encoding]::ASCII.GetByteCount($content) -eq $content.Length`
- **Propose changes before making them** -- do not edit code without confirmation

## Versioning and commits

This project uses [Semantic Versioning](https://semver.org) and [Conventional Commits](https://www.conventionalcommits.org).

**Commit message format:**
```
type(scope): short description
```

| Type | When |
|------|------|
| `feat` | New feature |
| `fix` | Bug fix |
| `docs` | Documentation only |
| `test` | Test additions or changes |
| `refactor` | Code change with no behaviour change |
| `chore` | Tooling, gitignore, config |

| Scope | Files / area |
|-------|-------------|
| `ps1` | Sync-Receipts.ps1 (general; use a narrower scope when one applies) |
| `logging` | Write-Log function and console output |
| `categories` | Categories feature: Get-Categories, Sync-CategorySheet, Categories.json |
| `accounts` | Accounts feature: Get-ValidAccounts, Accounts.xlsx, Accounts.template.xlsx |
| `sync-month` | Sync-Month function |
| `parse-receipt` | Parse-Receipt function |
| `xml` | Set-SubcategoryValidationXml XML patching |
| `tests` | tests/Sync-Receipts.Tests.ps1 |
| `readme` | README.md |
| `claude.md` | CLAUDE.md |
| `ci` | GitHub Actions workflows (.github/) |
| `config` | config.bat, .vscode/, batch launcher files |
| `changelog` | CHANGELOG.md |

**Versioning:**
- `MAJOR` -- breaking changes (file format changes, removed parameters, renamed files)
- `MINOR` -- new features
- `PATCH` -- bug fixes, docs, tests
- Version is tracked in the `Sync-Receipts.ps1` header and git tags (`v0.5.0`)
- `CHANGELOG.md` is updated manually when tagging a release

## Architecture

```
config.bat                <- local machine settings (gitignored); sets RECEIPTS_ROOT
config.template.bat       <- generic template committed to git
Accounts.template.xlsx    <- copy to RECEIPTS_ROOT\Accounts.xlsx and fill in accounts
Categories.json           <- category/subcategory definitions; committed, edit directly
Sync-Receipts.ps1         <- core automation (Excel COM)
Run-SyncReceipts.bat      <- launcher: calls config.bat, syncs current month
Run-SyncAllReceipts.bat   <- launcher: calls config.bat, syncs all months (-All)
Kill-Excel.bat            <- standalone utility: force-closes hung EXCEL.EXE processes
Deploy-SyncReceipts.bat   <- machine-specific deployment helper (gitignored)
```

The script files live in their own directory. The data (per-year workbooks and receipt folders) lives at `RECEIPTS_ROOT`, which is set in `config.bat`. Each year gets its own workbook (`2026.xlsx`, `2025.xlsx`, etc.) created automatically on first sync. `Categories.json` lives in the script directory and is read from `$PSScriptRoot`. The two locations are completely independent -- `-ReceiptsRoot` must always be provided explicitly; the script's own folder has no special meaning at runtime.

### Key functions in Sync-Receipts.ps1

| Function | Purpose |
|----------|---------|
| `Write-Log` | Writes timestamped, tagged log lines to the console; routes VERB-tagged messages to `Write-Verbose` |
| `Parse-Receipt` | Regex-parses a receipt filename stem into date, vendor, amount, method, account |
| `Read-PreservedCategoryValues` | Pure helper: extracts Category/Subcategory keyed by File Name from a 2D string array (no COM dependency; unit-testable) |
| `Get-ValidAccounts` | Reads 4-digit account numbers from `Accounts.xlsx` in `ReceiptsRoot`; falls back to Account sheet |
| `Get-Categories` | Reads category/subcategory data from `Categories.json` in `$PSScriptRoot` |
| `Sync-CategorySheet` | Writes category data from hashtable into the Category sheet (creates if absent, overwrites if present, hides the sheet) |
| `Get-ExcelColumnLetter` | Converts a 1-based column index to an Excel column letter (e.g. 1 -> "A", 27 -> "AA"); used to build named range address strings without COM |
| `Set-CategoryNamedRanges` | Creates named ranges in the workbook for Category/Subcategory dropdowns |
| `Set-SubcategoryValidationXml` | Post-save XML patch: injects both dropdown validations and fixes zip headers |
| `Set-MonthSheetOrder` | Sorts all month sheet tabs (4-digit YYMM names) into chronological order |
| `Sync-Month` | Main workhorse: creates/overwrites a month sheet and writes all receipt rows |

### Excel COM patterns used

- `New-Object -ComObject Excel.Application` -- headless Excel instance
- `$sheet.ListObjects.Add(...)` -- creates a structured table
- `$sheet.Hyperlinks.Add(...)` -- file hyperlinks in the File Name column
- `$workbook.Names.Add(name, ref)` -- creates/replaces named ranges
- Post-save XML patch via `System.IO.Compression.ZipFile` + binary header fix (see `Set-SubcategoryValidationXml`)
