# Claude Code Context — Sync-Receipts

## What this project is

A PowerShell script (`Sync-Receipts.ps1`) that uses Excel COM automation to parse receipt filenames and sync metadata into a formatted Excel workbook. See README.md for full usage details.

## Coding rules

- **Always add error handling and debug output** — every new block needs `try/catch` and `Write-Host` logging
- **Never use `$variable:` in double-quoted strings** — PowerShell interprets the colon as a drive separator; use `${variable}:` instead
- **No smart quotes or em-dashes** — the file must be pure ASCII. Non-ASCII characters break PowerShell parsing on the network share. Verify after any edit: `[System.Text.Encoding]::ASCII.GetByteCount($content) -eq $content.Length`
- **Propose changes before making them** — do not edit code without confirmation

## Architecture

```
config.bat              ← local machine settings (gitignored)
config.template.bat     ← generic template committed to git
Sync-Receipts.ps1       ← core automation (Excel COM)
Run-SyncReceipts.bat    ← launcher: calls config.bat, runs script for current month
Run-SyncAllReceipts.bat ← launcher: runs script with -All flag (no config.bat needed)
Deploy-SyncReceipts.bat ← machine-specific deployment helper (gitignored)
```

### Key functions in Sync-Receipts.ps1

| Function | Purpose |
|----------|---------|
| `Parse-Receipt` | Regex-parses a receipt filename stem into date, vendor, amount, method, account |
| `Get-ValidAccounts` | Reads 4-digit account numbers from the Account sheet |
| `Get-Categories` | Reads category/subcategory data from the Category sheet |
| `Set-CategoryNamedRanges` | Creates named ranges in the workbook for dropdown validation |
| `Sync-Month` | Main workhorse: creates/overwrites a month sheet and writes all receipt rows |

### Excel COM patterns used

- `New-Object -ComObject Excel.Application` — headless Excel instance
- `$sheet.ListObjects.Add(...)` — creates a structured table
- `$sheet.Hyperlinks.Add(...)` — file hyperlinks in the File Name column
- `$range.Validation.Add(3, 1, 1, "=Category")` — data validation dropdown via named range
- `$workbook.Names.Add(name, ref)` — creates/replaces named ranges

## Known issues

See [ISSUES.md](ISSUES.md) for the full breakdown of the subcategory dropdown problem and the binary zip-patching approach that needs to be implemented.
