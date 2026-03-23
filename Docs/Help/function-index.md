# Function Index

Quick reference for all key functions in Sync-Receipts.ps1. Functions marked with [OK] are unit-testable (pure; no COM dependency).

## Logging & Utility

| Function | Purpose | Testable |
|----------|---------|----------|
| `Write-SyncLog` | Writes timestamped, tagged log lines to console; routes VERB-tagged messages to `Write-Verbose` | -- |

## Receipt Data Parsing

| Function | Purpose | Testable |
|----------|---------|----------|
| `ConvertFrom-ReceiptFileName` | Regex-parses receipt filename stem into date, vendor, amount, method, account; two-pass parse validates method tokens | [OK] |
| `Get-ReceiptFlag` | Evaluates a parsed receipt row against configured rules; returns flag string; called by `Write-MonthSheet` for every receipt | [OK] |

## Configuration Loading

| Function | Purpose | Testable |
|----------|---------|----------|
| `Get-Methods` | Loads configurable non-Cash method tokens from `Config/Methods.json`; falls back to built-in defaults on missing/malformed file | [OK] |
| `Get-ValidAccounts` | Reads account records (Last4, Method, Holder, Institution, Account, Status) from `Accounts.xlsx` in `Config\`; skips if absent | -- |
| `Get-Categories` | Reads category/subcategory data from `Categories.json` in `Config/` | -- |

## Category Management

| Function | Purpose | Testable |
|----------|---------|----------|
| `Read-PreservedCategoryValues` | Extracts Category/Subcategory keyed by File Name from 2D string array; used to preserve user selections on re-sync | [OK] |
| `Test-SubcategoryValid` | Returns true if a subcategory belongs to a given category; used to clear stale subcategories on re-sync | [OK] |
| `Write-CategorySheet` | Writes category data from hashtable into Category sheet (creates if absent, overwrites if present, hides sheet) | -- |
| `Set-CategoryNamedRanges` | Creates named ranges in workbook for Category/Subcategory dropdowns | -- |
| `ConvertTo-ExcelRangeName` | Sanitizes category display name into valid Excel named range (e.g., "Food & Dining" -> "Food___Dining") | [OK] |

## Excel Utilities

| Function | Purpose | Testable |
|----------|---------|----------|
| `Get-ExcelColumnLetter` | Converts 1-based column index to Excel column letter (1 -> "A", 27 -> "AA"); builds named range addresses without COM | [OK] |

## Workbook Operations

| Function | Purpose | Testable |
|----------|---------|----------|
| `Write-MonthSheet` | Main workhorse: creates/overwrites month sheet, writes all receipt rows, applies formattin | -- |
| `Set-MonthSheetOrder` | Sorts all month sheet tabs (4-digit YYMM names) into chronological order | -- |
| `Set-SubcategoryValidationXml` | Post-save XML patch: injects dropdown validations and fixes zip headers; **must escape special characters** | -- |

## Test Files

- `Tests/<FunctionName>.Tests.ps1` -- Pester unit tests for pure functions
- `Tests/Lint.Tests.ps1` -- PSScriptAnalyzer validation
- `Tests/Integration/Invoke-SyncReceiptsTest.ps1` -- Syncs fixture; asserts Flag column

## Adding New Functions

1. **Pure logic?** -> Add a unit test in `Tests/<FunctionName>.Tests.ps1`
2. **Excel COM logic?** -> Integration tests only
3. **New method tokens, categories, or accounts?** -> Update corresponding `Get-*` function
4. **Error handling?** -> Always add `try/catch` + `Write-SyncLog` calls (see [coding-rules.md](coding-rules.md))

## Related

- [ARCHITECTURE-QUICK-REF.md](architecture-quick-ref.md) -- File structure and Excel COM patterns
- [CODING-RULES.md](coding-rules.md) -- Mandatory coding constraints
