# Testing Strategy

Unit testing, integration testing, Pester patterns, and CI behavior.

## Quick Rules

1. **Pure PowerShell functions -> unit test required** (marked [OK] in function-index.md)
2. **Excel COM functions -> integration test only** (no unit test)
3. **All tests must pass** before push (pre-push hook)
4. **CI mirrors local testing** (tests.yml runs same suite on github-latest)

---

## Unit vs. Integration Testing

### Unit Tests (Pure Functions)

**Characteristics:**
- No Excel COM dependencies
- Deterministic (same input -> same output)
- Run in microseconds
- No external files needed (or use fixtures)

**Functions requiring unit tests** ([OK] in function-index):
- `ConvertFrom-ReceiptFileName` -- regex parsing
- `Get-ReceiptFlag` -- validation logic
- `Get-Methods` -- config loading
- `Read-PreservedCategoryValues` -- array filtering
- `Test-SubcategoryValid` -- lookup logic
- `ConvertTo-ExcelRangeName` -- string sanitization
- `Get-ExcelColumnLetter` -- math logic

**Test location:**
```
Tests/<FunctionName>.Tests.ps1
```

**Example:**
```powershell
# Tests/ConvertFrom-ReceiptFileName.Tests.ps1
Describe "ConvertFrom-ReceiptFileName" {
    Context "Valid receipt filename" {
        It "parses date correctly" {
            $result = ConvertFrom-ReceiptFileName -FileName "260322_Amazon_49.99_CC_1234.pdf"
            $result.Date | Should -Be (Get-Date -Year 2026 -Month 3 -Day 22)
        }
    }
}
```

### Integration Tests (COM Functions)

**Characteristics:**
- Depend on Excel COM automation
- Require real workbooks or fixtures
- Run in seconds/minutes
- Local-only (not in CI)

**Functions with integration tests:**
- `Write-MonthSheet` -- writes to actual Excel
- `Write-CategorySheet` -- manipulates workbook structure
- `Set-SubcategoryValidationXml` -- post-save patching
- `Sync-Receipts.ps1` (main script)

**Test location:**
```
Tests/Integration/
  Invoke-SyncReceiptsTest.ps1
  Invoke-NewAccountsTemplateTest.ps1
  Invoke-InitializeSyncReceiptsTest.ps1
  Fixture/                        (test data)
    2026/2601/                    (receipt files for month 2601)
    Config/                       (fixture accounts.xlsx, categories.json, etc.)
```

**Why local-only:**
- Requires Windows + Excel installed
- CI (github-latest) may not have Excel
- Slow (Excel startup time)
- Test fixtures must be committed (2MB+ workbooks)

**CI testing strategy:**
- `tests.yml` runs `Tests/Lint.Tests.ps1` (PSScriptAnalyzer only)
- `pre-push` hook runs full Pester (unit + integration) locally
- Developers must verify integration tests pass before push

---

## Pester Test Structure

Standard Pester 5.x format:

```powershell
Describe "FunctionName" {
    BeforeEach {
        # Setup (runs before each It block)
    }

    Context "Scenario or feature" {
        It "should do X when Y" {
            $result = FunctionName -Param1 "value"
            $result | Should -Be "expected"
        }

        It "should handle empty input" {
            $result = FunctionName -Param1 ""
            $result | Should -BeNullOrEmpty
        }
    }

    Context "Another scenario" {
        It "should fail gracefully" {
            { FunctionName -Param1 $null } | Should -Throw
        }
    }

    AfterEach {
        # Cleanup (runs after each It block)
    }
}
```

### Common Assertions

| Assertion | Use For |
|-----------|---------|
| `Should -Be` | Exact equality (scalar) |
| `Should -BeExactly` | Case-sensitive string equality |
| `Should -BeNullOrEmpty` | Null or empty string/array |
| `Should -Contain` | Array contains value |
| `Should -Match` | Regex match |
| `Should -Throw` | Command throws exception |
| `Should -BeOfType` | Verify type |

### Example: ConvertFrom-ReceiptFileName.Tests.ps1

```powershell
Describe "ConvertFrom-ReceiptFileName" {
    Context "Valid filenames" {
        It "parses simple receipt" {
            $result = ConvertFrom-ReceiptFileName -FileName "260322_Amazon_49.99_CC_1234.pdf"
            $result.Vendor | Should -Be "Amazon"
            $result.Amount | Should -Be 49.99
            $result.Method | Should -Be "CC"
            $result.Account | Should -Be "1234"
        }

        It "handles vendor with spaces" {
            $result = ConvertFrom-ReceiptFileName -FileName "260322_Whole Foods_156.42_DB_5678.pdf"
            $result.Vendor | Should -Be "Whole Foods"
        }

        It "handles missing amount" {
            $result = ConvertFrom-ReceiptFileName -FileName "260228_Gift Card__VND_0000.pdf"
            $result.Amount | Should -Be 0
        }
    }

    Context "Invalid filenames" {
        It "throws on invalid date" {
            { ConvertFrom-ReceiptFileName -FileName "260230_Amazon_49.99_CC_1234.pdf" } | Should -Throw
        }

        It "handles unknown method" {
            $result = ConvertFrom-ReceiptFileName -FileName "260322_Amazon_49.99_XX_1234.pdf"
            $result.Method | Should -Be "Cash"  # Falls back to Cash
        }
    }
}
```

---

## Running Tests

### Locally

**Run all Pester tests:**
```powershell
Invoke-Pester -Path Tests/ -PassThru
```

**Run unit tests only:**
```powershell
Invoke-Pester -Path Tests/ -ExcludePath Tests/Integration/ -PassThru
```

**Run specific test file:**
```powershell
Invoke-Pester -Path Tests/ConvertFrom-ReceiptFileName.Tests.ps1 -PassThru
```

**Run integration tests:**
```powershell
cd Tests/Integration
.\Invoke-SyncReceiptsTest.ps1
```

### Using Batch Launcher

```bash
Tests\run-sync-test.bat
```

Runs full integration test suite with pre-commit check.

### Pre-Push Hook

When pushing to remote:
```bash
git push
```

Triggers `pre-push` hook which:
1. Runs linting (PSScriptAnalyzer)
2. Runs all Pester tests (unit + integration)
3. Blocks push if any test fails

---

## CI Pipeline

### tests.yml (GitHub Actions)

Runs on every push and pull request:

```yaml
name: Tests
on: [push, pull_request]

jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      - name: Run Pester
        run: Invoke-Pester -Path Tests/Lint.Tests.ps1 -PassThru
```

**Note:** Only runs `Lint.Tests.ps1` (PSScriptAnalyzer), not integration tests. Integration tests are local-only.

### commit-lint.yml

Enforces Conventional Commits format on PR titles.

### auto-release.yml

Watches for CHANGELOG entries and auto-tags on merge.

---

## Test Fixtures

Located in `Tests/Integration/Fixture/`.

### Receipt Files

Zero-byte placeholder files covering every flag scenario:

```
Fixture/
  2026/
    2601/                              (month 2601 = Jan 2026)
      260105_Costco_85.50_CC_1234.pdf   (valid)
      260110_Unknown_50_XX_1234.pdf     (unknown method)
      260115_Shell_42_CC_9999.pdf       (invalid account)
      260120_Amazon_0_CC_1234.pdf       (no amount)
      260125_Store_50_CC_0000.pdf       (closed account)
      260128_Vendor_30_DB_5678.pdf      (valid)
      260131_Item_100_VND_1234.pdf      (valid)
```

### Fixture Config

```
Fixture/Config/
  Accounts.xlsx               (fictitious accounts for testing)
  Categories.json             (test categories)
  Methods.json                (test methods)
```

---

## Writing a Unit Test for a New Function

1. **Create test file:**
   ```powershell
   # Tests/MyNewFunction.Tests.ps1
   Describe "MyNewFunction" {
       # Test blocks here
   }
   ```

2. **Import the function:**
   ```powershell
   BeforeAll {
       . (Join-Path (Split-Path $PSScriptRoot -Parent) "Scripts\Sync-Receipts.ps1")
   }
   ```

3. **Write test cases:**
   ```powershell
   Context "Happy path" {
       It "should return expected value" {
           $result = MyNewFunction -Param "value"
           $result | Should -Be "expected"
       }
   }
   ```

4. **Run test:**
   ```powershell
   Invoke-Pester -Path Tests/MyNewFunction.Tests.ps1
   ```

---

## Error Handling in Tests

**Test that errors are thrown correctly:**

```powershell
It "throws when input is invalid" {
    { MyFunction -BadParam } | Should -Throw "expected error message"
}
```

**Test that function handles errors gracefully:**

```powershell
It "continues after catching error" {
    $result = MyFunction -InvalidFile
    $result | Should -Be $null
}
```

---

## Debugging Test Failures

**Print debug info:**
```powershell
It "should process receipts" {
    $result = Get-ReceiptFlag -ParsedReceipt $receipt
    Write-Host "Result: $($result | ConvertTo-Json)" # Debug output
    $result | Should -Be ""
}
```

**Run single test:**
```powershell
Invoke-Pester -Path Tests/ConvertFrom-ReceiptFileName.Tests.ps1 -Filter "should parse simple receipt"
```

**Run with verbose output:**
```powershell
Invoke-Pester -Path Tests/ -Verbose
```

---

## Related

- [function-index.md](function-index.md) -- Which functions need tests ([OK] mark)
- [coding-rules.md](coding-rules.md) -- Error handling patterns for testable code
- [architecture-quick-ref.md](architecture-quick-ref.md) -- Test file locations
- [workflow.md](workflow.md) -- When to commit tests (same commit as new function)
