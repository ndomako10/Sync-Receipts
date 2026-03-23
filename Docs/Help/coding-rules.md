# PowerShell Coding Rules

Mandatory rules for all code contributions to Sync-Receipts.

## Error Handling & Logging

**Always add error handling and debug output** -- every new block needs `try/catch` and `Write-SyncLog` calls.

```powershell
try {
    # Your code here
    Write-SyncLog -Message "Processing started" -Tag VERB
} catch {
    Write-SyncLog -Message "Error: $_" -Tag ERROR
    throw
}
```

Tags:
- `-Tag WARN` -- warnings
- `-Tag ERROR` -- errors
- `-Tag VERB` -- diagnostic detail (routed to `Write-Verbose`)

## String Handling

### Never use `$variable:` in double-quoted strings

**BAD:**
```powershell
"The path is $variable:path"  # PowerShell treats `:` as a drive separator
```

**GOOD:**
```powershell
"The path is ${variable}:path"  # Use ${} to isolate the variable name
```

### ASCII-only files required

**No smart quotes or em-dashes** -- the file must be pure ASCII. Non-ASCII characters break PowerShell parsing on network shares.

Verify after edits:
```powershell
[System.Text.Encoding]::ASCII.GetByteCount($content) -eq $content.Length
```

## Date Handling

**Use .NET ParseExact format strings** -- write the actual format string, not informal names.

**CORRECT format strings:**
- `yyMMdd` = 2-digit year, month, day (e.g., `260323`)
- `yyyy` = 4-digit year
- `MM` = 2-digit month
- `dd` = 2-digit day

In documentation, always state the actual format string and explain each token:
> Date format: `yyMMdd` where `yy` = 2-digit year, `MM` = month (01-12), `dd` = day (01-31)

## XML Handling

**XML-escape string literals injected into XML** when building XML strings for the post-save patch (`Set-SubcategoryValidationXml`).

Escape:
- `&` -> `&amp;`
- `<` -> `&lt;`
- `>` -> `&gt;`

A bare `&` in injected XML causes Excel to report a parse error on open.

**Add a Pester assertion on the escaped form** whenever a new string literal is injected.

```powershell
# Example: injecting a category name into Excel validation XML
$safeName = $categoryName -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
$xmlFragment = "<data>$safeName</data>"

# In tests:
$xmlFragment | Should -Not -Match '(?<!amp);(?!#?\w+;)'  # No unescaped ampersand
```

## Testing Requirements

**Write Pester tests for new pure-PowerShell functions** -- functions with no COM dependency must have unit tests.

- One test file per function: `Tests/<FunctionName>.Tests.ps1`
- Pure helpers (with no Excel COM calls) must be testable
- Example: `Read-PreservedCategoryValues` is unit-testable; `Write-MonthSheet` is COM-dependent

Prefer pure helpers over COM-coupled logic wherever testability allows.

## Related

- See [ARCHITECTURE-QUICK-REM.md](architecture-quick-ref.md) for Excel COM patterns
- See [FUNCTION-INDEX.md](function-index.md) for existing helper functions
