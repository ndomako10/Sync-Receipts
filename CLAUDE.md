# Claude Code Context -- Sync-Receipts

## What This Project Is

A PowerShell automation tool for syncing receipt file metadata into Excel workbooks, with tests, setup scripts, and batch launchers. Uses Excel COM automation to parse receipt filenames and write formatted per-year workbooks. See README.md for full usage details.

## Workflow

> Issue tracking, commit discipline, documentation updates, and the propose-before-edit rule are governed by the [global CLAUDE.md](~/.claude/CLAUDE.md). This section documents project-specific overrides only.

### Commits
- **CHANGELOG.md entries are hand-crafted** -- write the entry at the top of `CHANGELOG.md` before tagging. Use user-facing language grouped under `### Added`, `### Changed (breaking)`, `### Fixed`, `### Removed`. The release workflow reads this entry and uses it as the GitHub Release body automatically.

## Coding Rules

- **Always add error handling and debug output** -- every new block needs `try/catch` and `Write-SyncLog` calls (use `-Tag WARN` for warnings, `-Tag ERROR` for errors, `-Tag VERB` for diagnostic detail)
- **Never use `$variable:` in double-quoted strings** -- PowerShell interprets the colon as a drive separator; use `${variable}:` instead
- **No smart quotes or em-dashes** -- the file must be pure ASCII. Non-ASCII characters break PowerShell parsing on the network share. Verify after any edit: `[System.Text.Encoding]::ASCII.GetByteCount($content) -eq $content.Length`
- **Use .NET ParseExact format strings for dates** -- write `yyMMdd`, not informal `YYMMDD`. In documentation, always use the actual format string and note what each token means (`yy` = 2-digit year, `MM` = month, `dd` = day)
- **XML-escape string literals injected into XML** -- when building XML strings in PowerShell for the post-save patch (`Set-SubcategoryValidationXml`), escape `&` as `&amp;`, `<` as `&lt;`, and `>` as `&gt;`. A bare `&` in injected XML causes Excel to report a parse error on open. Add a Pester assertion on the escaped form whenever a new string literal is injected.
- **Write Pester tests for new pure-PowerShell functions** -- functions with no COM dependency must have unit tests in `Tests/<FunctionName>.Tests.ps1` (one file per function). Prefer pure helpers (like `Read-PreservedCategoryValues`) over COM-coupled logic wherever testability allows

## Versioning and Commits

This project uses [Semantic Versioning](https://semver.org) and [Conventional Commits](https://www.conventionalcommits.org).

**Commit message format:** `type(scope): short description`

**Commit types:** `feat`, `fix`, `docs`, `test`, `refactor`, `chore`, `ci`, `perf`, `style`, `build` -- see global CLAUDE.md for definitions.

| Scope | Files / area |
|-------|-------------|
| `ps1` | Sync-Receipts.ps1 (general; use a narrower scope when one applies) |
| `logging` | Write-SyncLog function and console output |
| `categories` | Categories feature: Get-Categories, Write-CategorySheet, Config/Categories.json, Config/Templates/Categories.template.json |
| `accounts` | Accounts feature: Get-ValidAccounts, Config/Accounts.xlsx, Config/Templates/Accounts.template.xlsx, Scripts/New-AccountsTemplate.ps1 |
| `methods` | Configurable method tokens: Get-Methods, Config/Methods.json, Config/Templates/Methods.template.json |
| `write-month` | Write-MonthSheet function |
| `convert-receipt` | ConvertFrom-ReceiptFileName function |
| `xml` | Set-SubcategoryValidationXml XML patching |
| `tests` | Tests/ (per-function test files and Lint.Tests.ps1) |
| `adr` | Docs/ADRs/ |
| `readme` | README.md |
| `claude.md` | CLAUDE.md |
| `ci` | GitHub Actions workflows (.github/) |
| `config` | Config.ini, Config/ subfolder, .vscode/, batch launcher files |
| `changelog` | CHANGELOG.md |
| `contributing` | CONTRIBUTING.md |
| `security` | SECURITY.md |

**Versioning:**
- `MAJOR` -- breaking changes (file format changes, removed parameters, renamed files)
- `MINOR` -- new features
- `PATCH` -- bug fixes, docs, tests
- Version is tracked in git tags only (e.g. `v1.0.0`); the script header carries no version number
- `CHANGELOG.md` is hand-crafted; the release workflow finds the entry matching the tag and publishes it as the GitHub Release body

## Architecture

```
Setup.bat                 <- one-time setup launcher (runs Scripts\Initialize-SyncReceipts.ps1)
Config/
    Config.ini               <- local machine settings (gitignored); sets RECEIPTS_ROOT
    Accounts.xlsx            <- gitignored; personal accounts
    Categories.json          <- gitignored; personal categories
    Methods.json             <- gitignored; personal method token list
    Templates/
        Config.template.ini      <- generic template committed to git
        Accounts.template.xlsx   <- default accounts template; committed to git
        Categories.template.json <- default categories template; committed to git
        Methods.template.json    <- default method token list; committed to git
Scripts/
    Initialize-SyncReceipts.ps1 <- one-time setup: checks prerequisites, creates Config\Config.ini,
                             copies template files from Config\Templates\ to Config\, installs git hooks, creates shortcuts in RECEIPTS_ROOT
    Sync-Receipts.ps1        <- core automation (Excel COM)
    New-AccountsTemplate.ps1 <- regenerates Config\Templates\Accounts.template.xlsx from current schema; run after schema changes
    Install-GitHooks.ps1     <- copies Scripts\hooks\ into .git\hooks\
    hooks/
        commit-msg                  <- enforces Conventional Commits format (delegates to Invoke-CommitMsgCheck.ps1)
        Invoke-CommitMsgCheck.ps1   <- PowerShell implementation of commit-msg hook
        pre-commit                  <- ASCII check and PSScriptAnalyzer lint on staged .ps1 files (delegates to Invoke-PreCommitCheck.ps1)
        Invoke-PreCommitCheck.ps1   <- PowerShell implementation of pre-commit hook
        pre-push                    <- full Pester suite (delegates to Invoke-PrePushCheck.ps1)
        Invoke-PrePushCheck.ps1     <- PowerShell implementation of pre-push hook
Launchers/
    Run-SyncReceipts.bat     <- reads Config/Config.ini, syncs current month
    Run-SyncMonthReceipts.bat <- reads Config/Config.ini, syncs a specific month
    Run-SyncYearReceipts.bat  <- reads Config/Config.ini, syncs all months in a specific year
    Run-SyncAllReceipts.bat  <- reads Config/Config.ini, syncs all months (-All)
Docs/
    ADRs/                    <- Architecture Decision Records
        README.md            <- index of all ADRs
Tests/
    <Function>.Tests.ps1     <- Pester unit tests (one file per pure-PowerShell function)
    Lint.Tests.ps1           <- PSScriptAnalyzer validation
    run-sync-test.bat        <- local integration test launcher (tracked; not gitignored)
    Integration/             <- COM-dependent integration tests (local-only; not run in CI)
        Fixture/             <- zero-byte placeholder receipt files covering every flag scenario
            2026/2601/       <- month folder matching YearMonth=2601
        Config/              <- fixture config: Accounts.xlsx (fictitious), Categories.json, Methods.json
        Invoke-SyncReceiptsTest.ps1           <- syncs fixture; asserts Flag column (7 rows)
        Invoke-NewAccountsTemplateTest.ps1    <- runs New-AccountsTemplate.ps1; validates schema
        Invoke-InitializeSyncReceiptsTest.ps1 <- idempotency smoke test; asserts 4 .lnk shortcuts
.github/
    workflows/tests.yml          <- CI: runs Pester on windows-latest
    workflows/release.yml        <- publishes GitHub Release from CHANGELOG entry on tag push
    workflows/commit-lint.yml    <- enforces Conventional Commits format on pull requests
    workflows/labeler.yml        <- auto-applies labels to PRs based on changed files
    ISSUE_TEMPLATE/              <- bug report and feature request templates
    PULL_REQUEST_TEMPLATE.md     <- PR checklist template
    labeling.yml                 <- file-to-label mapping for the labeler workflow
    dependabot.yml               <- weekly dependency update checks
Kill-Excel.bat            <- standalone utility: force-closes hung EXCEL.EXE (gitignored)
CONTRIBUTING.md           <- dev guide: prerequisites, test instructions, commit format
CHANGELOG.md              <- version history; hand-crafted before each tag; release workflow reads the top entry as the GitHub Release body
SECURITY.md               <- responsible disclosure policy
.vscode/                  <- editor settings (extensions.json, settings.json, launch.json, tasks.json)
```

The script files live in their own directory. The data (per-year workbooks and receipt folders) lives at `RECEIPTS_ROOT`, which is set in `Config/Config.ini`. Each year gets its own workbook (`2026.xlsx`, `2025.xlsx`, etc.) created automatically on first sync. Workbooks are written to `WORKBOOKS_ROOT` when set in `Config/Config.ini`, defaulting to `RECEIPTS_ROOT`; this allows workbooks to be stored on a fast local drive while receipts remain on a network share (see ADR-006). `Categories.json` and `Accounts.xlsx` both live in `Config/` (gitignored) and are read via `Join-Path (Split-Path $PSScriptRoot -Parent) "Config"`. The two locations are completely independent -- `-ReceiptsRoot` defaults to the parent of Scripts/ but should always be set explicitly via Config/Config.ini.

### Key Functions in Sync-Receipts.ps1

| Function | Purpose |
|----------|---------|
| `Write-SyncLog` | Writes timestamped, tagged log lines to the console; routes VERB-tagged messages to `Write-Verbose` |
| `ConvertFrom-ReceiptFileName` | Regex-parses a receipt filename stem into date, vendor, amount, method, account; two-pass parse validates method tokens against the Methods list |
| `Read-PreservedCategoryValues` | Pure helper: extracts Category/Subcategory keyed by File Name from a 2D string array (no COM dependency; unit-testable) |
| `Get-Methods` | Loads configurable non-Cash method tokens from `Config/Methods.json`; falls back to built-in defaults on missing or malformed file |
| `Get-ValidAccounts` | Reads structured account records (Last4, Method, Holder, Institution, Account, Status) from `Accounts.xlsx` in `Config\`; skips validation if absent |
| `Get-Categories` | Reads category/subcategory data from `Categories.json` in `Config/` (repo root + "Config") |
| `Write-CategorySheet` | Writes category data from hashtable into the Category sheet (creates if absent, overwrites if present, hides the sheet) |
| `Get-ExcelColumnLetter` | Converts a 1-based column index to an Excel column letter (e.g. 1 -> "A", 27 -> "AA"); used to build named range address strings without COM |
| `ConvertTo-ExcelRangeName` | Sanitizes a category display name into a valid Excel named range identifier (e.g. "Food & Dining" -> "Food___Dining") |
| `Test-SubcategoryValid` | Pure helper: returns true if a subcategory belongs to a given category; used to clear stale subcategories on re-sync |
| `Set-CategoryNamedRanges` | Creates named ranges in the workbook for Category/Subcategory dropdowns |
| `Set-SubcategoryValidationXml` | Post-save XML patch: injects both dropdown validations and fixes zip headers |
| `Get-ReceiptFlag` | Evaluates a parsed receipt row against configured rules and returns a flag string; called by `Write-MonthSheet` for every receipt |
| `Set-MonthSheetOrder` | Sorts all month sheet tabs (4-digit YYMM names) into chronological order |
| `Write-MonthSheet` | Main workhorse: creates/overwrites a month sheet and writes all receipt rows |

### Excel COM Patterns Used

- `New-Object -ComObject Excel.Application` -- headless Excel instance
- `$sheet.ListObjects.Add(...)` -- creates a structured table
- `$sheet.Hyperlinks.Add(...)` -- file hyperlinks in the File Name column
- `$workbook.Names.Add(name, ref)` -- creates/replaces named ranges
- Post-save XML patch via `System.IO.Compression.ZipFile` + binary header fix (see `Set-SubcategoryValidationXml`)
