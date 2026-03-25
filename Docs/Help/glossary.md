# Glossary

Project terminology and abbreviations.

## Date & Time Formats

| Term | Format | Example | Meaning |
|------|--------|---------|---------|
| **yyMMdd** | 2-digit year, month, day | `260322` | Receipt date (March 22, 2026); used in filenames |
| **YYMM** | 2-digit year, month | `2601` | Month identifier (January 2026); used for sheet names and folder paths |
| **yyyy** | 4-digit year | `2026` | Full year |
| **MM** | 2-digit month | `03` | Month (01-12) |
| **dd** | 2-digit day | `22` | Day (01-31) |

## Account & Payment Terms

| Term | Definition | Example |
|------|-----------|---------|
| **Last4** | Final 4 digits of account number | `1234` in "****1234" |
| **Account Last4** | Full phrase for the account identifier field | Used in receipts and Accounts.xlsx |
| **Method** or **Payment Method** | Payment type token | CC (credit card), DB (debit), AP (ACH/transfer), Cash, VND (vendor/store credit) |
| **Status** | Account state field | Active, Closed, Inactive |
| **Method Token** | Short code for payment method | "CC", "DB", "AP", "VND", "CHK" (configured in Methods.json) |
| **Cash Fallback** | Default when method not found | If "XYZ" not in Methods.json, defaults to "Cash" and flags "Unknown Method" |

## File & Folder Organization

| Term | Path | Example |
|------|------|---------|
| **RECEIPTS_ROOT** | User-configured root folder | `C:\Users\YourName\Documents\Receipts` |
| **WORKBOOKS_ROOT** | Optional; where year-workbooks stored | `C:\Users\YourName\Documents\Workbooks` (defaults to RECEIPTS_ROOT) |
| **Receipt Folder Structure** | `RECEIPTS_ROOT/YYYY/YYMM/` | `C:\Receipts\2026\2601\` for Jan 2026 receipts |
| **Workbook Location** | `WORKBOOKS_ROOT/YYYY.xlsx` | `C:\Workbooks\2026.xlsx` for year 2026 |
| **Config Folder** | `Config/` (gitignored) | User's local configuration files |
| **Config/Templates/** | `Config/Templates/` (committed) | Default templates for setup |

## Excel Terminology

| Term | Definition | Purpose |
|------|-----------|---------|
| **Month Sheet** | One Excel sheet per month | Sheet name = YYMM (e.g., "2601") |
| **ListObject** | Excel structured table | Auto-creates headers, enables filtering |
| **Named Range** | Excel identifier for a range | Used for dropdown validation sources |
| **Data Validation** | Excel dropdown list | Category and Subcategory dropdowns |
| **INDIRECT Formula** | Excel formula that evaluates string | Used for dependent dropdowns (show subcategories based on category) |
| **XML Patch** | Post-save modification to .xlsx internals | Injects validation XML after save (using ZipFile) |

## Receipt Parsing & Validation

| Term | Definition | Example |
|------|-----------|---------|
| **Filename Stem** | Filename without extension | `260322_Amazon_49.99_CC_1234` from `260322_Amazon_49.99_CC_1234.pdf` |
| **Parsed Receipt** | Extracted components | `@{ Date=..., Vendor=Amazon, Amount=49.99, Method=CC, Account=1234 }` |
| **Two-Pass Validation** | Parsing in two steps | Pass 1: Extract from filename; Pass 2: Validate against config |
| **Flag** | Indicator for problematic receipts | "Unknown Method", "Account Not Found", "Account Closed", empty = OK |
| **Receipt Row** | One row in month sheet | Contains all parsed + user-entered fields (Date, Vendor, Amount, Category, Flag, etc.) |

## Features

| Term | Definition | Related Files |
|------|-----------|---------------|
| **Category Dropdown** | User selects category (Food & Dining, Transportation, etc.) | Categories.json, Get-Categories, Write-CategorySheet |
| **Subcategory Dropdown** | Dependent dropdown; shows subcats for selected category | Categories.json, INDIRECT formula, Set-SubcategoryValidationXml |
| **Account Validation** | Checks if Last4 exists in Accounts.xlsx and is Active | Accounts.xlsx, Get-ValidAccounts, Get-ReceiptFlag |
| **Method Validation** | Checks if payment method is in Methods.json | Methods.json, Get-Methods, ConvertFrom-ReceiptFileName |
| **Sensitive Data Detection** | Blocks commit if SSN, account number, credit card, etc. detected | SensitivePatterns.json, Invoke-SensitiveDataCheck, pre-commit hook |
| **Flag System** | Marks receipts with issues for manual review | Get-ReceiptFlag, Write-MonthSheet |
| **Hyperlinks** | Clickable "File Name" column that opens receipt file | Set-MonthSheet |

## Processing & Synchronization

| Term | Definition | Command |
|------|-----------|---------|
| **Sync** | Process receipts and update workbook | `Scripts\Sync-Receipts.ps1` |
| **Sync Current Month** | Sync the current month | `Run-SyncReceipts.bat` |
| **Sync Specific Month** | Sync a user-selected month | `Run-SyncMonthReceipts.bat` |
| **Sync Year** | Sync all months in a year | `Run-SyncYearReceipts.bat` |
| **Sync All** | Sync all months ever (all folders) | `Run-SyncAllReceipts.bat` |

## Logging & Debugging

| Term | Definition | Usage |
|------|-----------|-------|
| **Write-SyncLog** | Logging function | `Write-SyncLog "Message" -Tag VERB` |
| **Tag** | Log message category | VERB (verbose), WARN (warning), ERROR (error), or blank (info) |
| **Verbose Output** | Detailed diagnostic messages | Run with `-Verbose` flag; shows VERB-tagged messages |

## Git & Workflow

| Term | Definition | Scope |
|------|-----------|-------|
| **Conventional Commit** | Standardized commit message format | `type(scope): description` |
| **Type** | Commit classification | feat, fix, docs, test, refactor, chore, ci, perf, style, build |
| **Scope** | Subsystem being changed | ps1, logging, categories, accounts, methods, tests, etc. |
| **Pre-Commit Hook** | Local check before commit | ASCII validation, linting, sensitive data scan |
| **Pre-Push Hook** | Local check before push | Full Pester test suite |
| **[Unreleased]** | Section in CHANGELOG for pending features | Used for auto-release trigger |

## Configuration

| Term | Definition | Type |
|------|-----------|------|
| **Config.ini** | Machine-local settings (paths, etc.) | Text; gitignored |
| **Accounts.xlsx** | Table of authorized accounts | Excel; gitignored |
| **Categories.json** | Hierarchical category/subcategory structure | JSON; gitignored |
| **Methods.json** | Array of valid payment method tokens | JSON; gitignored |
| **SensitivePatterns.json** | Regex patterns for sensitive data detection | JSON; gitignored |
| **Template Files** | Default versions of above (in Config/Templates/) | All types; committed to git |

## Testing & CI

| Term | Definition |
|------|-----------|
| **Unit Test** | Test pure PowerShell function (no COM/Excel dependency) |
| **Integration Test** | Test Excel COM logic with real workbook |
| **Pester** | PowerShell testing framework (used for all tests) |
| **PSScriptAnalyzer** | PowerShell linting tool (enforces style rules) |
| **Fixture** | Test data files (zero-byte receipts, test config) |
| **CI** | Continuous Integration (GitHub Actions workflows) |

## Architecture & Design

| Term | Definition | Related |
|------|-----------|---------|
| **ADR** | Architecture Decision Record | Docs/ADRs/ |
| **ADR-006** | Design decision: separate WORKBOOKS_ROOT and RECEIPTS_ROOT | Allows fast local workbooks + network receipt storage |
| **Sheet Order** | Month sheets sorted chronologically (not creation order) | Set-MonthSheetOrder |
| **XML-Escape** | Convert `&` -> `&amp;`, `<` -> `&lt;`, `>` -> `&gt;` | Required for category names in validation XML |

## Related

- [context-selection.md](../context-selection.md) -- How to use this glossary
- [receipt-format.md](receipt-format.md) -- Detailed receipt syntax
- [config-schema.md](config-schema.md) -- Config file structures
