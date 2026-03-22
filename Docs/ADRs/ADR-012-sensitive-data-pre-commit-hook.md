# ADR-012: Sensitive Data Pre-Commit Hook Design

## Status
Accepted

## Context

ADR-009 was committed with real institution names and actual card numbers used as examples.
The error was caught manually. The project stores personal financial data in gitignored files,
but documentation, ADRs, and code could accidentally reference real account numbers or
institution names. A pre-commit check was needed to catch this class of mistake automatically.

Several design decisions had non-obvious tradeoffs.

## Decision

Add a sensitive data check to the pre-commit hook implemented as a separate
`Invoke-SensitiveDataCheck.ps1` file containing two pure, unit-testable functions:
`Invoke-SensitiveDataCheck` and `Get-AccountsLast4`. The check is wired into
`Invoke-PreCommitCheck.ps1`, which loads patterns from `Config\SensitivePatterns.json`
(falling back to the committed template) and builds dynamic patterns from the user's
`Config\Accounts.xlsx` at hook runtime.

## Alternatives considered

- **Fold the check into `Invoke-PreCommitCheck.ps1`** -- simpler file count, but
  `Invoke-PreCommitCheck.ps1` is a script with a main execution block, making it awkward
  to unit-test individual functions in isolation. A separate file with only function
  definitions can be dot-sourced cleanly from Pester tests.

- **Read `Accounts.xlsx` via Excel COM** -- consistent with how `Get-ValidAccounts` works in
  `Sync-Receipts.ps1`, but COM requires Excel to be installed and running, adds ~5 seconds
  of startup time to every commit, and cannot run in CI. Reading the xlsx as a ZIP archive
  and parsing the shared strings XML directly (as `Set-SubcategoryValidationXml` already
  does for writing) avoids all three problems.

- **Static patterns only (no Accounts.xlsx integration)** -- simpler, but the most important
  sensitive values -- the user's real Last4 account identifiers -- are only known at runtime
  from `Accounts.xlsx`. Static patterns cannot catch a bare `1234` unless the user manually
  adds every account number to the config, which they are unlikely to maintain.

- **Per-line comment suppression using `# sensitive-ok` or similar** -- `# nocheck` was
  chosen because it is shorter, already recognised in several lint ecosystems, and reads
  naturally as "do not check this line". A project-specific token would require more
  documentation to be discoverable.

- **Empty `fileTypes` meaning "apply to all files"** -- the implementation treats an empty
  `fileTypes` array as "apply to all files" (consistent with opt-in filtering). The
  `keyword-search` template entry uses `[".ps1", ".md"]` explicitly to prevent the template
  file from matching its own placeholder regex when staged.

## Consequences

- Every commit that stages `.ps1`, `.json`, or `.md` files runs the sensitive data check in
  addition to the existing ASCII and lint checks. On a typical commit this adds under 100 ms.
- `Config\SensitivePatterns.json` is gitignored and created by `Setup.bat`, consistent with
  `Categories.json` and `Methods.json`. The hook falls back to the committed template when
  the user copy is absent.
- False positives in documentation or fixture files can be suppressed inline with `# nocheck`
  or via the `allowlist` array in `SensitivePatterns.json`, without modifying the hook script.
- `Get-AccountsLast4` is a pure function testable with inline XML strings, so account-number
  detection has unit test coverage without requiring `Accounts.xlsx` to be present in CI.
