Review all Pester tests for correctness, coverage, and style.

Read each of the following files in full:
- Tests/ConvertFrom-ReceiptFileName.Tests.ps1
- Tests/ConvertTo-ExcelRangeName.Tests.ps1
- Tests/Get-Categories.Tests.ps1
- Tests/Get-ExcelColumnLetter.Tests.ps1
- Tests/Get-Methods.Tests.ps1
- Tests/Get-ReceiptFlag.Tests.ps1
- Tests/Initialize-SyncReceipts.Tests.ps1
- Tests/Lint.Tests.ps1
- Tests/Read-PreservedCategoryValues.Tests.ps1
- Tests/Set-SubcategoryValidationXml.Tests.ps1
- Tests/Test-SubcategoryValid.Tests.ps1
- Tests/Write-SyncLog.Tests.ps1

Also run:
- `gh issue list --state open --label test` to check for known outstanding gaps

Check each area and flag any issues:

1. **BDD naming** -- Every test must follow the Describe/Context/It convention from CLAUDE.md:
   - `Describe` names the unit under test (function or script name)
   - `Context` uses `given` / `when` / `with` language to describe the condition
   - `It` states the expected outcome as a plain-language sentence
   - Flag any `Context` block that does not use this language, and any `It` description that reads like an implementation log rather than a behavioural specification

2. **Coverage completeness** -- For each test file, cross-reference against the corresponding function in `Scripts/Sync-Receipts.ps1`. Flag any clearly missing cases:
   - Happy-path variants (different valid inputs)
   - Error / fallback branches (missing file, malformed input, empty input)
   - Edge cases called out in the function's `.DESCRIPTION` or known from the issue tracker

3. **Test hygiene**
   - Temp directories: all file I/O tests must use Pester's `$TestDrive` rather than `[System.IO.Path]::GetTempPath()` + `NewGuid()`. Flag any manual temp dir creation that is not cleaned up.
   - Isolation: each `It` block must be self-contained. Flag shared mutable state set up outside `BeforeAll` / `BeforeEach` that could bleed between tests.
   - Assertions: flag `It` blocks with no `Should` assertion, or assertions that only check `$true`/`$false` without also asserting specific values or error messages where those are meaningful.

4. **Lint.Tests.ps1 correctness** -- Verify:
   - All `.ps1` files in `Scripts/` are covered by a `Describing PSScriptAnalyzer --` block
   - The settings file path (`.config/PSScriptAnalyzerSettings.psd1`) is asserted to exist before `Invoke-ScriptAnalyzer` is called; flag if this guard is absent

5. **Open test issues** -- For each issue returned by `gh issue list --state open --label test`, state whether the gap it describes is present in the current tests or has already been addressed (flag for closing if so).

## Output format

For each finding: `SEVERITY | FILE:LINE | one-sentence problem | one-sentence fix`
Severity: CRITICAL (broken/data-loss), MAJOR (correctness gap), MINOR (polish).
If no issues in an area: `OK: <area>`
No narrative. No file contents. Findings only.
