# Features Overview

High-level explanation of each major feature and how it works.

## Categories Feature

### What It Does

Provides hierarchical category/subcategory dropdowns in Excel for organizing receipts.

### Example

Users select from:
- **Categories:** Food & Dining, Transportation, Utilities, etc.
- **Subcategories:** When "Food & Dining" is selected, user picks "Groceries", "Restaurants", "Delivery", etc.

### Implementation

1. **Config:** Categories stored in `Config/Categories.json`
   ```json
   {
     "Food & Dining": ["Groceries", "Restaurants", "Delivery"],
     "Transportation": ["Gas", "Parking", "Rideshare"]
   }
   ```

2. **Loading:** `Get-Categories` reads the JSON file
   - Falls back to empty hashtable if missing or malformed
   - Returns PowerShell hashtable: `@{ "Food & Dining" = @("Groceries", ...) }`

3. **Writing to Excel:**
   - `Write-CategorySheet` creates a hidden Category sheet
   - Lists all categories and subcategories in columns

4. **Dropdowns:**
   - `Set-CategoryNamedRanges` creates Excel named ranges
   - Category dropdown references Category column
   - Subcategory dropdown uses INDIRECT formula (dependent on Category selection)

5. **XML Injection:**
   - `Set-SubcategoryValidationXml` injects data validation XML after save
   - Allows dependent dropdowns to work ("when Category = X, show Subcategories for X")
   - **Note:** Must XML-escape special characters (`&` -> `&amp;`, etc.)

6. **Preservation on Re-Sync:**
   - `Read-PreservedCategoryValues` extracts user-selected categories
   - When re-syncing same month, user selections are preserved
   - Clears stale subcategories if categories changed

### Functions Involved

- `Get-Categories` -- loads config
- `Write-CategorySheet` -- creates sheet
- `Set-CategoryNamedRanges` -- creates dropdown ranges
- `Set-SubcategoryValidationXml` -- injects data validation
- `ConvertTo-ExcelRangeName` -- sanitizes names for Excel
- `Read-PreservedCategoryValues` -- preserves selections
- `Test-SubcategoryValid` -- validates membership

---

## Accounts Feature

### What It Does

Validates receipt accounts against an authorized account list and flags suspicious transactions.

### Example

Users define accounts in `Config/Accounts.xlsx`:
- Last4 = `1234`, Method = `CC`, Status = `Active` -> OK
- Last4 = `9999`, Method = `CC`, Status = `Closed` -> Flagged "Account Closed"

### Implementation

1. **Config:** Accounts stored in `Config/Accounts.xlsx`
   ```
   | Last4 | Method | Holder | Institution | Account | Status |
   |-------|--------|--------|-------------|---------|--------|
   | 1234  | CC     | John   | Chase       | Visa    | Active |
   | 5678  | DB     | John   | Bank of A   | Chk     | Active |
   | 9999  | CC     | John   | Amex        | Closed  | Closed |
   ```

2. **Loading:** `Get-ValidAccounts` reads the spreadsheet
   - Parses table rows into PowerShell objects
   - Returns array of account records with columns

3. **Validation in Receipt Parsing:**
   - `ConvertFrom-ReceiptFileName` validates Last4 against accounts
   - If Last4 not found, sets `AccountValid = $false`
   - Reads `Status` field to determine if account is active/closed/inactive

4. **Flagging:**
   - `Get-ReceiptFlag` checks account status
   - If status = "Closed", flags receipt "Account Closed"
   - If status = "Inactive", flags receipt "Account Inactive"
   - Empty status or "Active" = no flag

5. **Display:**
   - Receipt shows masked Last4 (`****1234`) in Excel
   - Pre-commit hook masks full account numbers from diffs

### Functions Involved

- `Get-ValidAccounts` -- loads accounts
- `ConvertFrom-ReceiptFileName` -- validates Last4
- `Get-ReceiptFlag` -- flags closed/inactive accounts

---

## Methods Feature

### What It Does

Defines valid payment method tokens (CC, DB, ACH, etc.) and validates receipts.

### Example

Default methods:
```json
["Cash", "CC", "DB", "AP", "VND"]
```

Users can customize in `Config/Methods.json`.

### Implementation

1. **Config:** Methods stored in `Config/Methods.json`
   ```json
   ["Cash", "CC", "DB", "AP", "VND", "CHK"]
   ```

2. **Loading:** `Get-Methods` reads the array
   - Falls back to built-in defaults if missing/malformed
   - Returns PowerShell array: `@("Cash", "CC", "DB", ...)`

3. **Validation in Receipt Parsing:**
   - `ConvertFrom-ReceiptFileName` looks up method token in list
   - Two-pass validation:
     - Pass 1: Extract method from filename
     - Pass 2: Check if method is in Methods list
   - If not found, defaults to "Cash" and stores original in `MethodOriginal`

4. **Usage:**
   - Receipts without recognized methods are silently converted to "Cash"
   - Flag system marks unknown methods for review
   - Prevents parsing errors from bad input

### Functions Involved

- `Get-Methods` -- loads methods
- `ConvertFrom-ReceiptFileName` -- validates tokens

---

## Sensitive Patterns Feature

### What It Does

Detects sensitive data (SSN, account numbers, etc.) in staged files before commit.

### Example

Patterns configured in `Config/SensitivePatterns.json`:
```json
{
  "SSN": "\\b\\d{3}-\\d{2}-\\d{4}\\b",
  "Account": "\\b[0-9]{10,17}\\b",
  "Phone": "\\b\\d{3}[-.]?\\d{3}[-.]?\\d{4}\\b"
}
```

### Implementation

1. **Config:** Patterns stored in `Config/SensitivePatterns.json`
   - User can add custom patterns
   - Uses regex (standard .NET format)

2. **Pre-Commit Hook:**
   - `pre-commit` hook runs `Invoke-SensitiveDataCheck`
   - Scans all staged files (not working tree)
   - If matches found, blocks commit with list of matches

3. **Scanning:**
   - `Invoke-SensitiveDataCheck` performs pattern scan
   - Returns array of matches + file + line + pattern name
   - User must remove data or add file to `.gitignore`

4. **Account Masking:**
   - `Get-AccountsLast4` reads accounts from Excel
   - Pre-commit pattern also detects Last4 from account list
   - Prevents accidental account number commits

5. **Post-Save Masking:**
   - After sync, account numbers can be masked in workbook backups
   - Controlled via SensitivePatterns.json

### Functions Involved

- `Invoke-SensitiveDataCheck` -- scans files
- `Get-AccountsLast4` -- extracts account numbers for pattern matching
- `pre-commit` hook -- blocks unsafe commits

---

## Flag System

### What It Does

Marks problematic receipts with flags in Excel for manual review.

### Table of Flags

| Flag | Meaning | Action |
|------|---------|--------|
| (empty) | No issues | Process normally |
| "Unknown Method" | Method not in Methods.json | Add method or fix filename |
| "Account Not Found" | Last4 not in Accounts.xlsx | Add account or correct filename |
| "Account Closed" | Account status is Closed | Verify if legitimate transaction |
| "Account Inactive" | Account status is Inactive | Verify if account should be active |
| "No Amount" | Amount is 0 or blank | Fill in actual amount |

### Implementation

1. **Parsing:**
   - `ConvertFrom-ReceiptFileName` extracts components
   - Validates against config files

2. **Evaluation:**
   - `Get-ReceiptFlag` checks all rules
   - Returns appropriate flag string (or empty if OK)
   - Called by `Write-MonthSheet` for every receipt row

3. **Display in Excel:**
   - Flag column shows message (default: empty)
   - Users quickly spot issues (flagged rows stand out)
   - Can sort/filter by Flag column

4. **Example:**
   ```
   Date       Vendor    Amount  Method  Account  Flag
   03/22/26   Amazon    $49.99  CC      ****1234
   03/23/26   Shell     $42.00  XX      ****1234 Unknown Method
   03/24/26   Best Buy  $150    CC      ****9999 Account Not Found
   03/25/26   Costco    $85.50  CC      ****5678
   ```

### Functions Involved

- `ConvertFrom-ReceiptFileName` -- validation
- `Get-ReceiptFlag` -- flag determination
- `Write-MonthSheet` -- displays flag in Excel

---

## Hyperlinks Feature

### What It Does

Links receipt files directly from Excel for quick access.

### Implementation

1. **File Organization:**
   - Receipts live in `RECEIPTS_ROOT/YYYY/YYMM/filename.pdf`
   - Month sheet references this path

2. **Hyperlink Creation:**
   - `Write-MonthSheet` calls `$sheet.Hyperlinks.Add(...)`
   - File Name column contains clickable links
   - Ctrl+Click opens the PDF (or file viewer)

3. **Relative Paths:**
   - Hyperlinks stored relative to workbook location
   - If workbook moves, links may break (design limitation)
   - See ADR-006 for separation of WORKBOOKS_ROOT and RECEIPTS_ROOT

---

## Month Sheet Organization

### What It Does

Automatically creates and sorts monthly tabs in the workbook.

### Implementation

1. **Month Sheets:**
   - One sheet per month (name = `YYMM`, e.g., `2601`)
   - `Write-MonthSheet` creates/overwrites sheets

2. **Sorting:**
   - `Set-MonthSheetOrder` sorts tabs chronologically
   - Prevents random tab order (default is creation order)
   - Uses 4-digit YYMM format for natural sort

3. **Year Workbooks:**
   - Each year gets its own workbook (2026.xlsx, 2025.xlsx, etc.)
   - Workbooks are auto-created on first sync
   - Format: StandardDate tab + per-month tabs + hidden Category sheet

---

## Related

- [receipt-format.md](receipt-format.md) -- Filename syntax and parsing
- [config-schema.md](config-schema.md) -- All configuration options
- [function-index.md](function-index.md) -- Functions implementing features
- [architecture-quick-ref.md](architecture-quick-ref.md) -- Excel COM patterns
