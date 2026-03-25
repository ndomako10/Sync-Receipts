# Debugging Guide

Troubleshooting tools, logging, common failures, and diagnostic workflows.

## Logging System: Write-SyncLog

All diagnostic output goes through `Write-SyncLog` function with tagged messages.

### Function Signature

```powershell
Write-SyncLog -Message <string> [-Tag <string>] [-Verbose]
```

### Tags

| Tag | Output | Usage | Example |
|-----|--------|-------|---------|
| (none) | Console | General info | `Write-SyncLog "Syncing month 2601"` |
| `VERB` | `Write-Verbose` | Diagnostic details | `Write-SyncLog "Loading config" -Tag VERB` |
| `WARN` | Console + Yellow | Warnings | `Write-SyncLog "Account not found" -Tag WARN` |
| `ERROR` | Console + Red | Errors | `Write-SyncLog "File not found: $_" -Tag ERROR` |

### Viewing Verbose Output

Enable verbose logging to see all VERB-tagged messages:

```powershell
# Run sync with verbose output
Scripts\Sync-Receipts.ps1 -All -Verbose

# Or enable for a session
$VerbosePreference = "Continue"
Scripts\Sync-Receipts.ps1 -All
```

### Example Log Output

```
[2026-03-23 14:30:15] Syncing receipts: C:\Receipts
[2026-03-23 14:30:15] VERB: Loading methods from config
[2026-03-23 14:30:16] Processing month 2601 (13 files)
[2026-03-23 14:30:17] WARN: Account not found (Last4: 9999)
[2026-03-23 14:30:18] Finished: 10 receipts written, 2 flagged
```

---

## Utilities

### Kill-Excel.bat

Force-closes hung Excel processes.

```bash
Kill-Excel.bat
```

**Use when:**
- Excel freezes during workbook write
- COM object won't release
- Multiple EXCEL.EXE processes running in Task Manager

**Does:**
```powershell
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
```

---

## Common Failures & Solutions

### File Lock / "File is in Use" Error

**Symptom:**
```
Error: The file "2026.xlsx" is in use by another process
```

**Causes:**
- Excel still has file open
- Another sync process running
- Antivirus scanning file

**Solutions:**

1. **Close Excel:**
   ```bash
   Kill-Excel.bat
   ```

2. **Check for running sync:**
   ```powershell
   Get-Process powershell | Where-Object { $_.MainWindowTitle -like "*Sync-Receipts*" }
   ```

3. **Check antivirus:**
   - Temporarily exclude workbook folder
   - Add folder to Windows Defender exclusions

4. **Retry after delay:**
   ```powershell
   Start-Sleep -Seconds 5
   Scripts\Sync-Receipts.ps1 -Month 2601
   ```

---

### "Config File Not Found"

**Symptom:**
```
Error: Cannot find Config\Config.ini
```

**Causes:**
- Setup not run (Initialize-SyncReceipts.ps1)
- Config.ini deleted
- Script running from wrong directory

**Solutions:**

1. **Run setup:**
   ```bash
   Setup.bat
   ```

2. **Check Config path:**
   ```powershell
   Join-Path (Split-Path $PSScriptRoot -Parent) "Config\Config.ini"
   # Should point to: C:\YourPath\Config\Config.ini
   ```

3. **Verify Config exists:**
   ```powershell
   Test-Path "C:\Users\YourName\OneDrive - The Peplinski Group\Workspace\Repos\Sync-Receipts\Config\Config.ini"
   ```

---

### "Invalid Receipt Filename"

**Symptom:**
```
WARN: Cannot parse receipt: invalid_filename.pdf
```

**Causes:**
- Filename doesn't match pattern: `YYMMDD_Vendor_Amount_Method_Account.ext`
- Underscores in vendor name (reserved delimiter)
- Invalid date (e.g., Feb 30)

**Solutions:**

1. **Check filename syntax** (see [receipt-format.md](receipt-format.md))
   - Must be: `260322_Amazon_49.99_CC_1234.pdf`
   - Not: `Best_Buy_receipt.pdf`

2. **Fix vendor name:**
   ```
   [X] Best_Buy_150_CC_1234.pdf    (underscore not allowed in vendor)
   [OK] BestBuy_150_CC_1234.pdf     (remove underscore)
   [OK] Best Buy_150_CC_1234.pdf    (space is OK)
   ```

3. **Validate date:**
   ```
   [X] 260230_Vendor_50_CC_1234.pdf    (Feb doesn't have 30 days)
   [OK] 260228_Vendor_50_CC_1234.pdf    (Feb 28 is valid)
   ```

---

### "Unknown Account" or "Account Closed" Flag

**Symptom:**
```
Receipt flagged: Account Not Found
or
Receipt flagged: Account Closed
```

**Causes:**
- Last4 digits not in Accounts.xlsx
- Account status in Excel is "Closed" or "Inactive"

**Solutions:**

1. **Check account exists:**
   - Open `Config\Accounts.xlsx`
   - Look for Last4 = 1234
   - If missing, add it

2. **Check account status:**
   ```
   Status = "Active"     -> OK
   Status = "Closed"     -> Flagged
   Status = "Inactive"   -> Flagged
   Status = ""           -> OK (blank = Active)
   ```

3. **Update status if needed:**
   - Open `Config\Accounts.xlsx`
   - Change "Closed" to "Active" if still in use
   - Save file

---

### "Unknown Method" Flag

**Symptom:**
```
Receipt flagged: Unknown Method
```

**Causes:**
- Method code not in Methods.json
- Typo in method (e.g., "Cc" instead of "CC")

**Solutions:**

1. **Check Methods.json:**
   ```powershell
   Get-Content Config\Methods.json | ConvertFrom-Json
   # Should show: ["Cash", "CC", "DB", "AP", ...]
   ```

2. **Verify method in receipt:**
   ```
   [X] 260322_Amazon_49.99_XYZ_1234.pdf    (XYZ not in methods)
   [OK] 260322_Amazon_49.99_CC_1234.pdf     (CC is in methods)
   ```

3. **Add method if needed:**
   - Edit `Config\Methods.json`
   - Add to array: `["Cash", "CC", "DB", "NewMethod"]`
   - Re-sync receipts

---

### "Sensitive Data Detected" (Commit Blocked)

**Symptom:**
```
Error: Sensitive data detected in staged files
  Pattern: SSN
  File: Config\Config.ini
  Match: 123-45-6789
```

**Causes:**
- Pre-commit hook detected SSN, account number, credit card, etc.
- Match in a file being committed

**Solutions:**

1. **Remove sensitive data:**
   ```bash
   git checkout HEAD -- Config\Config.ini  # Discard changes
   # Edit file and remove SSN
   git add Config\Config.ini
   git commit
   ```

2. **Add file to .gitignore:**
   ```bash
   echo "Config/SensitiveFile.txt" >> .gitignore
   git add .gitignore
   git commit -m "chore: ignore sensitive file"
   # File won't be tracked going forward
   ```

3. **Customize patterns:**
   - Edit `Config\SensitivePatterns.json`
   - Remove pattern or make less strict
   - Re-commit

---

### Excel Won't Open After Sync

**Symptom:**
```
"Excel cannot open the file. The file format may be corrupted."
```

**Causes:**
- XML injection failed (Set-SubcategoryValidationXml)
- Unescaped special characters in category names
- Corrupted workbook

**Solutions:**

1. **Check category names:**
   - Open `Config\Categories.json`
   - Look for unescaped `&`, `<`, `>` in names
   - XML will escape them, but excessive nesting can corrupt
   - Example: Remove `&` from "Food & Dining" -> "Food and Dining"

2. **Regenerate workbook:**
   ```powershell
   Remove-Item "C:\Workbooks\2026.xlsx"  # Delete corrupted file
   Scripts\Sync-Receipts.ps1 -Year 2026  # Re-create from scratch
   ```

3. **Check logs for errors:**
   ```powershell
   Scripts\Sync-Receipts.ps1 -All -Verbose
   # Look for ERROR-tagged messages
   ```

---

### Pre-Commit Hook Blocks Commit (Linting)

**Symptom:**
```
Error: PSScriptAnalyzer violation detected
  Rule: PSUseShouldProcessForStateChangingFunctions
  Line: 42
```

**Causes:**
- PowerShell code violates style rule
- ASCII validation failed (smart quotes, em-dashes)

**Solutions:**

1. **Fix style issue:**
   ```powershell
   # Run analyzer to see all violations
   Invoke-ScriptAnalyzer -Path Scripts\Sync-Receipts.ps1

   # Fix the issue in the file (line 42)
   # Then stage and retry
   git add Scripts\Sync-Receipts.ps1
   git commit -m "fix(ps1): ..."
   ```

2. **Check for non-ASCII characters:**
   ```powershell
   $content = Get-Content Scripts\Sync-Receipts.ps1 -Raw
   [System.Text.Encoding]::ASCII.GetByteCount($content) -eq $content.Length
   # Should return $true
   ```

3. **Remove smart quotes:**
   - Find all `"` and `"` (curly quotes) and replace with `"` (straight)
   - Find all `'` and `'` (curly quotes) and replace with `'` (straight)
   - Find all `--` (em-dash) and replace with `--` (double dash)

---

### Pre-Push Hook Fails (Tests)

**Symptom:**
```
Error: Test failed
  [X] ConvertFrom-ReceiptFileName should parse simple receipt
```

**Causes:**
- Unit test assertion failed
- Code bug introduced
- Test fixture outdated

**Solutions:**

1. **Run failing test locally:**
   ```powershell
   Invoke-Pester -Path Tests\ConvertFrom-ReceiptFileName.Tests.ps1 -Verbose
   ```

2. **Debug the test:**
   ```powershell
   # Run just the failing case
   Invoke-Pester -Path Tests\ConvertFrom-ReceiptFileName.Tests.ps1 `
     -Filter "should parse simple receipt"
   ```

3. **Fix the code or test:**
   - Examine function
   - Check test assertion
   - Fix bug or update test
   - Commit: `git commit -m "fix(tests): ..."`
   - Push: `git push`

---

## Debug Workflow Example

### Scenario: Receipts Not Syncing

```powershell
# 1. Enable verbose logging
$VerbosePreference = "Continue"

# 2. Run with specific month
Scripts\Sync-Receipts.ps1 -Month 2601

# 3. Check output for VERB messages and errors
# Look for:
#   - Config loading errors
#   - File not found errors
#   - Parse failures
#   - Flag messages

# 4. If config error, verify paths
Test-Path "C:\Path\To\Config\Config.ini"
Test-Path "C:\Path\To\Receipts\2026\2601"

# 5. If parse error, check filename format
Get-ChildItem "C:\Path\To\Receipts\2026\2601" -Filter "*.pdf"
# Should be: YYMMDD_Vendor_Amount_Method_Account.pdf

# 6. If still stuck, check workbook
Test-Path "C:\Path\To\Workbooks\2026.xlsx"

# 7. Try regenerating workbook
Remove-Item "C:\Path\To\Workbooks\2026.xlsx"
Scripts\Sync-Receipts.ps1 -Year 2026
```

---

## Related

- [coding-rules.md](coding-rules.md) -- Error handling patterns with Write-SyncLog
- [testing-strategy.md](testing-strategy.md) -- Running tests locally
- [github-workflows.md](github-workflows.md) -- Pre-commit hook details
- [receipt-format.md](receipt-format.md) -- Valid filename syntax
