# Receipt Format & Parsing

Receipt filename syntax, parsing logic, flag rules, and validation.

## Filename Syntax

Receipt filenames follow a strict pattern for Excel sync:

```
yyMMdd Vendor $Amount [Method [Account]].extension
```

**Format:** Space-delimited; Method and Account are optional.

### Components

| Component | Format | Example | Purpose | Notes |
|-----------|--------|---------|---------|-------|
| **Date** | `yyMMdd` | `260322` | Receipt date | yy=2-digit year, MM=month, dd=day; can use separators (e.g., `yy-MM-dd`) with -DateFormat parameter |
| **Vendor** | Text | `Amazon`, `Whole Foods` | Merchant name | No length limit; spaces are OK |
| **Amount** | `$` + decimal | `$49.99`, `-$20.00` | Receipt total | Dollar sign required; supports negative amounts (e.g., refunds) |
| **Method** | Single word | `Card`, `Cash`, `Check` | Payment type | Optional; if omitted, defaults to "". Validated against configured Methods list |
| **Account** | 4 chars | `1234`, `xxxx`, `----` | Account identifier | Optional if Method omitted; **required if Method present and not Cash**. Shown as masked (****1234) in Excel |
| **Extension** | File type | `.pdf`, `.jpg` | Document type | Not parsed; kept as-is |

### Valid Examples

```
260322 Amazon $49.99 Card 1234.pdf
260101 Whole Foods $156.42 Card 5678.jpg
260215 Shell $42.00 Card 1234.pdf
260228 Gift Card $0.00 Card.pdf
260315 Check $500.00 Check 0000.pdf
260320 Cash Withdrawal $12.50 Cash.pdf
260101 Target -$20.00 Card 1234.pdf    (negative amount / refund)
260316 Starbucks $5.27                  (no method or account; will be flagged)
```

### Invalid Examples

```
260322Amazon$49.99Card1234.pdf          # Missing spaces (delimiters required)
020322 Amazon $49.99 Card 1234.pdf      # Year is 02 (likely typo; current year ~26)
260322 Amazon $49.99 InvalidMethod 1234.pdf    # Method not in Methods list (will fail)
260322 Amazon $49.99 Card.pdf           # Non-Cash method without account number
260322 Amazon 49.99 Card 1234.pdf       # Missing $ symbol
260322 Amazon $49.99.00 Card 1234.pdf   # More than 2 decimal places
BADFILENAME.pdf                         # Does not match pattern
```

---

## Parsing: ConvertFrom-ReceiptFileName

Function: `ConvertFrom-ReceiptFileName` in Sync-Receipts.ps1

### Two-Pass Validation

**Pass 1: Broad Structure Check**
- Matches filename against broad regex: `yyMMdd Vendor $Amount [Method [Account]]`
- Validates date format (`yyMMdd`) is valid calendar date
- Confirms amount has dollar sign and exactly 2 decimal places: `-?\$[\d]+\.[\d]{2}`
- Accepts any single word as Method (or omits if no method provided)
- Validates account is 4 digits, "xxxx", "----", or omitted

**Pass 2: Method Validation**
- If Method is present: looks up in configured Methods list
- "Cash" is always valid (built-in, no need to configure)
- Any other method not in Methods list causes parse error: "Unrecognised method"
- If Method is not "Cash" and Account is missing, causes parse error: "Could not parse filename"

### Return Value

PowerShell hashtable:

```powershell
@{
    OK           = [bool]      # $true if parse succeeded, $false otherwise
    ParseError   = [string]    # Error reason if OK is $false ("Unrecognised method", "Could not parse filename", etc.)
    Date         = [datetime]  # Parsed date as DateTime object; $null if OK is $false
    Vendor       = [string]    # Vendor name (e.g., "Amazon")
    Amount       = [string]    # Numeric amount string with optional minus (e.g., "49.99" or "-20.00")
    Method       = [string]    # Method token (e.g., "Card", "Cash", "Check") or "" if omitted
    Account      = [string]    # Account identifier ("1234", "xxxx", "----") or "" if omitted or Cash
}
```

### Usage Example

```powershell
$result = ConvertFrom-ReceiptFileName -Stem "260322 Amazon $49.99 Card 1234"
$result.OK        # $true
$result.Date      # 3/22/2026
$result.Vendor    # Amazon
$result.Amount    # 49.99
$result.Method    # Card
$result.Account   # 1234

# With no method
$result2 = ConvertFrom-ReceiptFileName -Stem "260322 Starbucks $5.27"
$result2.OK       # $true
$result2.Method   # ""
$result2.Account  # ""

# With error
$result3 = ConvertFrom-ReceiptFileName -Stem "260322 Amazon $49.99 Card"  # Missing account for non-Cash
$result3.OK        # $false
$result3.ParseError # "Could not parse filename"
```

---

## Flag System

Flags are assigned during sync based on receipt parsing and account validation. Flagged receipts appear in the Excel Flag column for manual review.

### Flag Sources

| Condition | Flag | Resolution |
|-----------|------|------------|
| Parse error (method not in Methods list, account missing for non-Cash, invalid date, etc.) | "Parse error: [reason]" | Fix filename to match: `yyMMdd Vendor $Amount [Method [Account]]` |
| Method present but not in configured Methods list | "Unrecognised method: [method]" | Add method to Config/Methods.json or correct filename |
| Non-Cash method without account number | "Parse error: Could not parse filename" | Add account number (4 digits) to filename |
| Method omitted (optional) | "Method missing" | Add method and account to filename, or leave blank if intentional |
| Account not found in Accounts.xlsx | "Account not found: [Last4]" | Add account to Accounts.xlsx or correct Last4 in filename |
| Account status in Accounts.xlsx is "Inactive" | "Account inactive: [Last4]" | Verify if account should be active; update Accounts.xlsx status |
| Amount is 0 (zero dollars) | "Zero amount" | Confirm zero amount is intentional (e.g., gift card placeholder) |
| All validations pass | (empty) | No action needed |

### Example Output

```
Date       Vendor    Amount  Method  Account  Flag
03/22/26   Amazon    $49.99  Card    ****1234 (empty)
01/02/26   Walmart   -$15.00 Card    ****9999 Account not found: 9999
01/04/26   Shell     $50.00  Gas     (empty)  Unrecognised method: Gas
01/05/26   Costco    $80.00  Card    (empty)  Parse error: Could not parse filename
```

---

## Validation Rules

### Date (`yyMMdd`)

- Must be a valid calendar date (e.g., `260230` is invalid -- Feb doesn't have 30 days)
- Year range: typically 20-30 for years 2020-2030 (parseable with 2-digit token `yy`)
- Can use separators if specified with -DateFormat (e.g., `yy-MM-dd` for `26-03-22`)
- Default format is `yyMMdd` (no separators)

### Vendor Name

- Can be any text
- Spaces are OK (e.g., "Whole Foods", "Best Buy")
- No length limit (will truncate in Excel if too long for cell)
- Special characters OK; will be XML-escaped on save

### Amount

- **Required**: Dollar sign (`$`)
- Format: `-?\$[\d]+\.[\d]{2}` (optional minus, dollar sign, digits, period, exactly 2 decimal places)
- Examples: `$49.99`, `-$20.00` (refund), `$0.00`
- Must have exactly 2 decimal places

### Method

- **Optional**: If omitted, parse succeeds but flags "Method missing"
- Must match a token in configured Methods list (default: `Card`, `Check`, `Checking`, `Savings`, `Transfer`, `Wire`)
- "Cash" is always valid (built-in; no need to configure)
- Case-sensitive (e.g., "Card" [OK], "card" [X])
- Single word only; spaces not allowed
- If not found in Methods list, parse fails: "Unrecognised method"

### Account

- **Optional if Method omitted**: Returns as ""
- **Required if Method is non-Cash**: Parse fails if missing
- Format: 4 characters -- either 4 digits (e.g., `1234`), "xxxx", or "----"
- If not found in Accounts.xlsx, flags "Account not found: [Last4]"
- If found but status is "Inactive", flags "Account inactive: [Last4]"

---

## Folder Structure

Receipts are organized in a year/month hierarchy:

```
RECEIPTS_ROOT/
  YYYY/              (4-digit year)
    YYMM/            (2-digit year + month)
      receipt1.pdf
      receipt2.pdf
      ...
```

Example:
```
C:\Receipts\
  2026/
    2601/            (January 2026)
      260105 Costco $85.50 Card 1234.pdf
      260115 Shell $42.00 Card 1234.pdf
      ...
    2602/            (February 2026)
      260205 Amazon $49.99 Card 1234.pdf
      ...
    2603/            (March 2026)
      ...
```

**Month sheets** in the Excel workbook match the folder names (e.g., sheet "2601" for `RECEIPTS_ROOT/2026/2601/`).

**WORKBOOKS_ROOT** (optional): If set in Config.ini, year-workbooks (2026.xlsx, 2025.xlsx, etc.) are written to a separate location, allowing fast local storage while receipts remain on a network share.

---

## Example Processing

### Scenario 1: Valid Receipt

Given filename: `260322 Amazon $49.99 Card 1234.pdf`

```
1. Parse:
   Date:     260322 -> 3/22/2026
   Vendor:   Amazon
   Amount:   $49.99 -> "49.99" (stored as string)
   Method:   Card
   Account:  1234

2. Pass 1 validation (broad structure):
   [OK] Date is valid calendar date
   [OK] Vendor is present
   [OK] Amount has $ and 2 decimals
   [OK] Method is a single word
   [OK] Account is 4 digits

3. Pass 2 validation (method check):
   [OK] "Card" is in Methods list
   [OK] Account is provided (required for non-Cash)

4. Account lookup:
   [OK] 1234 found in Accounts.xlsx, Status = "Active"

5. Excel output:
   Date:          03/22/26
   Vendor:        Amazon
   Amount:        $49.99
   Method:        Card
   Account:       ****1234
   Category:      (user selects)
   Subcategory:   (user selects)
   Flag:          (empty)
```

### Scenario 2: Missing Account (Non-Cash)

Given filename: `260315 Shell $50.00 Card.pdf`

```
1. Parse:
   Date:     260315
   Vendor:   Shell
   Amount:   $50.00
   Method:   Card
   Account:  (missing)

2. Pass 2 validation:
   [X] "Card" method requires account (non-Cash methods must have account)
   ParseError: "Could not parse filename"

3. Result:
   OK = $false
   File skipped with warning
```

### Scenario 3: Unknown Method

Given filename: `260315 Shell $50.00 Gas 1234.pdf`

```
1. Pass 1 validation:
   [OK] All structural checks pass

2. Pass 2 validation:
   [X] "Gas" not in Methods list
   ParseError: "Unrecognised method"

3. Result:
   OK = $false
   File skipped with warning
```

### Scenario 4: Refund (Negative Amount)

Given filename: `260101 Target -$20.00 Card 1234.pdf`

```
1. Parse:
   Date:     260101
   Vendor:   Target
   Amount:   -$20.00 -> "-20.00"
   Method:   Card
   Account:  1234

2. All validations pass, stored as negative value

3. Excel output:
   Amount:  -$20.00
   Flag:    (empty)
```

### Scenario 5: Cash Payment (No Account)

Given filename: `260320 Coffee Shop $5.50 Cash.pdf`

```
1. Parse:
   Date:     260320
   Vendor:   Coffee Shop
   Amount:   $5.50
   Method:   Cash
   Account:  (omitted, OK for Cash)

2. Pass 2 validation:
   [OK] "Cash" is always valid
   [OK] Account not required for Cash

3. Excel output:
   Amount:    $5.50
   Method:    Cash
   Account:   (empty)
   Flag:      (empty)
```

---

## Related

- [config-schema.md](config-schema.md) -- Methods.json and Accounts.xlsx structure
- [coding-rules.md](coding-rules.md) -- Date format strings (yyMMdd)
- [function-index.md](function-index.md) -- ConvertFrom-ReceiptFileName function
- [features-overview.md](features-overview.md) -- Flag system and validation
- [glossary.md](glossary.md) -- YYMM, yyMMdd, Last4, and other terminology
