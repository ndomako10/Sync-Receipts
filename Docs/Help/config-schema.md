# Configuration Schema Reference

All configuration files, their structure, templates, and usage.

## Quick Reference

| File | Type | Location | Template | Purpose | Gitignored |
|------|------|----------|----------|---------|-----------|
| Config.ini | Text | `Config/` | Config.template.ini | Local machine settings | [OK] |
| Accounts.xlsx | Excel | `Config/` | Accounts.template.xlsx | Account validation records | [OK] |
| Categories.json | JSON | `Config/` | Categories.template.json | Category/subcategory hierarchies | [OK] |
| Methods.json | JSON | `Config/` | Methods.template.json | Payment method tokens | [OK] |
| SensitivePatterns.json | JSON | `Config/` | SensitivePatterns.template.json | Regex patterns for sensitive data detection | [OK] |

All templates are in `Config/Templates/` and committed to git.

## Config.ini

Local machine settings. Created by `Initialize-SyncReceipts.ps1` from `Config.template.ini`.

### Required Settings

```ini
RECEIPTS_ROOT=C:\Users\YourName\Documents\Receipts
```

Sets the root folder where receipt files are organized by year and month (e.g., `RECEIPTS_ROOT\2026\2601\`).

### Optional Settings

```ini
WORKBOOKS_ROOT=C:\Users\YourName\Documents\Workbooks
```

Target folder for year-based Excel workbooks (2026.xlsx, 2025.xlsx, etc.). If omitted, defaults to `RECEIPTS_ROOT`. Allows storing workbooks on a fast local drive while receipts remain on a network share (see ADR-006).

### Initialization

One-time setup:
```bash
Setup.bat
# or manually
Scripts\Initialize-SyncReceipts.ps1
```

This creates `Config/Config.ini` by copying `Config/Templates/Config.template.ini` and prompts for RECEIPTS_ROOT.

---

## Categories.json

Category and subcategory hierarchies for Excel dropdown validation.

### Structure

```json
{
  "Food & Dining": [
    "Groceries",
    "Restaurants",
    "Delivery",
    "Coffee & Snacks"
  ],
  "Transportation": [
    "Gas",
    "Parking",
    "Public Transit",
    "Rideshare",
    "Car Maintenance"
  ],
  "Utilities": [
    "Electric",
    "Gas",
    "Water",
    "Internet"
  ]
}
```

### Rules

- Top level = Category (e.g., "Food & Dining")
- Second level = Array of Subcategories (e.g., ["Groceries", "Restaurants"])
- All string values (`&` symbols allowed; will be XML-escaped on save)
- Flat structure only (no nesting deeper than 2 levels)

### How It's Used

1. `Get-Categories` loads the file
2. `Write-CategorySheet` creates the Category + Subcategory sheet
3. `Set-CategoryNamedRanges` creates Excel named ranges for dropdown validation
4. `Set-SubcategoryValidationXml` injects XML for dependent dropdowns
5. On re-sync: `Read-PreservedCategoryValues` preserves user-selected categories

### Initialization

Copied from `Config/Templates/Categories.template.json` during Setup. Edit to add custom categories.

---

## Accounts.xlsx

Structured account record table for validation. Created from template by `New-AccountsTemplate.ps1`.

### Columns

| Column | Type | Purpose | Validation |
|--------|------|---------|-----------|
| **Last4** | Text | Final 4 digits of account number | Used to match receipt parsing |
| **Method** | Text | Payment method (CC, DB, AP, VND, Cash) | Must exist in Methods.json (except Cash) |
| **Holder** | Text | Account owner name | Informational |
| **Institution** | Text | Bank/Credit card issuer | Informational |
| **Account** | Text | Full account identifier or nickname | Informational |
| **Status** | Text | Active, Closed, Inactive, etc. | Flags inactive accounts |

### Example

| Last4 | Method | Holder | Institution | Account | Status |
|-------|--------|--------|-------------|---------|--------|
| 1234 | CC | John Doe | Chase | Personal Visa | Active |
| 5678 | DB | John Doe | Bank of America | Checking | Active |
| 9012 | AP | John Doe | American Express | Business | Closed |

### How It's Used

1. `Get-ValidAccounts` reads the file
2. Receipt filename parsing validates account Last4 against this table
3. Account status is checked to flag closed accounts
4. Used by pre-commit hook to mask account data

### Initialization

First time, copy template:
```bash
Setup.bat  # Creates Config/Accounts.xlsx from template
```

After schema changes, regenerate:
```bash
Scripts\New-AccountsTemplate.ps1  # Preserves existing data, updates schema
```

---

## Methods.json

Configurable payment method tokens (non-Cash methods).

### Structure

```json
[
  "CC",    // Credit Card
  "DB",    // Debit
  "AP",    // ACH/Bank Transfer
  "VND",   // Vendor (store credit)
  "CHK"    // Check
]
```

### Rules

- Simple JSON array of strings
- No whitespace within tokens (used in regex)
- Case matters (e.g., "CC" vs "cc")
- Any length (1-10 chars typical)
- Used by `ConvertFrom-ReceiptFileName` in two-pass validation
- If a receipt doesn't match a listed method, it defaults to "Cash"

### How It's Used

1. `Get-Methods` loads the array
2. `ConvertFrom-ReceiptFileName` validates method tokens against this list
3. Falls back to built-in defaults if file is missing or malformed
4. Used during receipt parsing (two-pass validation)

### Initialization

Copied from `Config/Templates/Methods.template.json` during Setup. Edit to add custom method tokens.

---

## SensitivePatterns.json

Regex patterns for detecting sensitive data (SSN, account numbers, etc.) before commit.

### Structure

```json
{
  "SSN": "\\b\\d{3}-\\d{2}-\\d{4}\\b",
  "Account Number": "\\b[0-9]{10,17}\\b",
  "Routing Number": "\\b\\d{9}\\b",
  "Credit Card": "\\b[0-9]{13,19}\\b",
  "Phone": "\\b\\d{3}[-.]?\\d{3}[-.]?\\d{4}\\b"
}
```

### Rules

- Object of pattern names -> regex strings
- Regexes are literal strings (backslashes are literal in JSON)
- Used by `Invoke-SensitiveDataCheck` in the `pre-commit` hook
- Matches are reported; commit is blocked if sensitive data found
- Users can add custom patterns (names, medical info, etc.)

### How It's Used

1. `Invoke-SensitiveDataCheck` loads the patterns
2. Pre-commit hook scans staged files against patterns
3. If matches found, hook blocks commit with error message
4. User must remove sensitive data or .gitignore the file

### Initialization

Copied from `Config/Templates/SensitivePatterns.template.json` during Setup.

---

## Template Initialization Flow

```
1. User runs Setup.bat
   |
   [|-----> Scripts/Initialize-SyncReceipts.ps1
   |   [|-----> Checks PowerShell prerequisites
   |   [|-----> Prompts for RECEIPTS_ROOT
   |   [|-----> Creates Config/Config.ini from Config.template.ini
   |   [|-----> Copies Config.template.ini -> Config/Config.ini (overwrites if exists)
   |   [|-----> Copies all templates from Config/Templates/ to Config/
   |   |   [|-----> Accounts.template.xlsx -> Accounts.xlsx
   |   |   [|-----> Categories.template.json -> Categories.json
   |   |   [|-----> Methods.template.json -> Methods.json
   |   |   [`-----> SensitivePatterns.template.json -> SensitivePatterns.json
   |   [|-----> Scripts/Install-GitHooks.ps1 (installs pre-commit, commit-msg, pre-push)
   |   [`-----> Creates desktop shortcuts in RECEIPTS_ROOT

2. All user edits happen in Config/ (not Config/Templates/)
   Config/Config.ini is gitignored
   Config/Categories.json is gitignored
   Config/Accounts.xlsx is gitignored
   etc.

3. Templates in Config/Templates/ remain committed to git
   for future setups and schema reference.
```

## When to Regenerate Templates

### New-AccountsTemplate.ps1

Run this **after changing the Accounts.xlsx schema**:

```bash
Scripts\New-AccountsTemplate.ps1
```

This:
- Reads current Accounts.xlsx (preserves user data)
- Regenerates Config/Templates/Accounts.template.xlsx with new schema
- Does not overwrite user's Config/Accounts.xlsx

Use case: Adding a new column (e.g., "CreditLimit") to the account table.

### Other Templates

For Categories.json, Methods.json, SensitivePatterns.json:
- Edit the files in `Config/` directly (gitignored)
- Commit template changes separately to `Config/Templates/` if adding defaults

---

## Related

- [architecture-quick-ref.md](architecture-quick-ref.md) -- Data locations and file structure
- [workflow.md](workflow.md) -- Setup workflow (Initialize-SyncReceipts.ps1)
- [testing-strategy.md](testing-strategy.md) -- Configuration in test fixtures
- [features-overview.md](features-overview.md) -- How each config is used
