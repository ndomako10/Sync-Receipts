# Architecture Quick Reference

File structure and key patterns. See CLAUDE.md for detailed architecture.

## Directory Structure

```
Scripts/
  [|---- Sync-Receipts.ps1           (core automation; main entry point)
  [|---- Initialize-SyncReceipts.ps1  (one-time setup)
  [|---- New-AccountsTemplate.ps1     (regenerates accounts template)
  [|---- Install-GitHooks.ps1         (installs git hooks)
  [`---- hooks/
      [|---- commit-msg (enforces Conventional Commits)
      [|---- pre-commit (ASCII check, lint, sensitive data scan)
      [`---- pre-push (runs full Pester suite)

Config/
  [|---- Config.ini                   (local settings; gitignored; sets RECEIPTS_ROOT, WORKBOOKS_ROOT)
  [|---- Accounts.xlsx                (personal accounts; gitignored)
  [|---- Categories.json              (personal categories; gitignored)
  [|---- Methods.json                 (method tokens; gitignored)
  [|---- SensitivePatterns.json        (sensitive data patterns; gitignored)
  [`---- Templates/
      [|---- Config.template.ini
      [|---- Accounts.template.xlsx
      [|---- Categories.template.json
      [|---- Methods.template.json
      [`---- SensitivePatterns.template.json

Launchers/
  [|---- Run-SyncReceipts.bat        (syncs current month)
  [|---- Run-SyncMonthReceipts.bat   (syncs specific month)
  [|---- Run-SyncYearReceipts.bat    (syncs all months in year)
  [`---- Run-SyncAllReceipts.bat     (syncs all months)

Tests/
  [|---- <FunctionName>.Tests.ps1    (Pester unit tests; one per pure function)
  [|---- Lint.Tests.ps1               (PSScriptAnalyzer)
  [|---- run-sync-test.bat            (integration test launcher)
  [`---- Integration/
      [|---- Invoke-SyncReceiptsTest.ps1
      [|---- Invoke-NewAccountsTemplateTest.ps1
      [`---- Invoke-InitializeSyncReceiptsTest.ps1

Docs/
  [`---- ADRs/                        (Architecture Decision Records)
```

## Data Locations

| Item | Location | Notes |
|------|----------|-------|
| **RECEIPTS_ROOT** | `Config/Config.ini` | Source folder for receipt files; set per machine |
| **WORKBOOKS_ROOT** | `Config/Config.ini` | Target for year-workbooks; defaults to RECEIPTS_ROOT |
| **Workbooks** | `WORKBOOKS_ROOT/` | One workbook per year (2026.xlsx, 2025.xlsx, etc.) |
| **Receipt folders** | `RECEIPTS_ROOT/YYYY/YYMM/` | Month-based folder structure; month sheets reference this |
| **Categories** | `Config/Categories.json` | User data; Category and Subcategory dropdowns |
| **Accounts** | `Config/Accounts.xlsx` | User data; validates account Last4 in receipts |
| **Methods** | `Config/Methods.json` | User data; validated method tokens (non-Cash) |
| **Sensitive patterns** | `Config/SensitivePatterns.json` | User data; patterns to detect before commit |

**Default paths:** Config files are loaded via `Join-Path (Split-Path $PSScriptRoot -Parent) "Config"` -- this works correctly whether Scripts/ is under repo root or elsewhere, as long as Config/ is at the parent level.

## Excel COM Patterns

Standard patterns used throughout Sync-Receipts.ps1:

```powershell
# Create a headless Excel instance
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Create/open workbook and get worksheet
$workbook = $excel.Workbooks.Open($path)
$sheet = $workbook.Worksheets.Item($name)

# Create a structured table (ListObject)
$listObject = $sheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $range)

# Create hyperlinks (e.g., File Name column)
$sheet.Hyperlinks.Add($cell, $target) | Out-Null

# Create/replace named ranges (for dropdowns)
$workbook.Names.Add($name, $reference) | Out-Null

# Post-save XML patch (injects validation XML and fixes zip headers)
# See Set-SubcategoryValidationXml for details
```

## Key Configuration Files

**Config.ini** -- Machine-local settings (template: `Config.template.ini`)
```ini
RECEIPTS_ROOT=C:\Path\To\Receipts
WORKBOOKS_ROOT=C:\Path\To\Workbooks (optional; defaults to RECEIPTS_ROOT)
```

**Categories.json** -- Category/Subcategory structure (template: `Categories.template.json`)
```json
{
  "Food & Dining": ["Groceries", "Restaurants"],
  "Transportation": ["Gas", "Parking", "Rideshare"]
}
```

**Methods.json** -- Non-Cash payment method tokens (template: `Methods.template.json`)
```json
["CC", "DB", "AP", "VND"]
```

**Accounts.xlsx** -- Structured account records with columns: Last4, Method, Holder, Institution, Account, Status (template: `Accounts.template.xlsx`)

**SensitivePatterns.json** -- Custom regex patterns for sensitive data detection (template: `SensitivePatterns.template.json`)

## Related

- [FUNCTION-INDEX.md](function-index.md) -- All functions, test status, and dependencies
- [CODING-RULES.md](coding-rules.md) -- Error handling, string escaping, date formats
- [COMMIT-SCOPES.md](commit-scopes.md) -- Scope reference for files listed here
- [ADR-006](ADRs/README.md) -- Workbook storage strategy (local vs. network)
