## [2.0.0] - 2026-03-20

### Added
- Add configurable date format with per-field parse error flags, closes #25
- Add DATE_FORMAT to Config.template.bat and launchers, refs #25
- Allow Method and Account to be omitted; flag Method missing rows, closes #28
- Add WorkbooksRoot parameter to separate workbook output from receipts root, closes #31
- Prompt for WORKBOOKS_ROOT during setup, defaulting to RECEIPTS_ROOT
- Replace Config.bat with Config.env for safe editing, closes #37

### Documentation
- Add Releasing section and contributing scope, refs #30
- Update changelog and versioning workflow for git-cliff, refs #30
- Add vendor description, fix amount sign convention
- Fix amount sign convention in Write-MonthSheet doc
- Standardize date format string to yyMMdd, closes #29
- Standardize date format string to yyMMdd, refs #29
- Document -DateFormat parameter and parse error flag values, refs #25
- Document optional Method/Account, Method missing flag, and fix examples, refs #28
- Document new launchers, first-run sync all, and day-to-day vs correction use
- Document WorkbooksRoot parameter and WORKBOOKS_ROOT config variable, refs #31
- Document M and d token separator requirement in DateFormat, refs #32
- Document WORKBOOKS_ROOT prompt and folder creation confirmation in setup steps
- Add CODE_OF_CONDUCT.md
- Remove workflow rules now covered by global ~/.claude/CLAUDE.md
- Update Accounts.xlsx and Categories.json paths for Config\ move
- Update architecture and function table for Config\ file moves
- Update accounts and categories scope descriptions
- Fix Config.bat path, add missing launchers, add Docs/ADRs
- Add six retroactive ADRs for core architectural decisions
- Fix heading capitalisation to match project title-case style
- Update all Config.bat references to Config.env
- Add ADR-007 keep workbooks as xlsx not xlsm
- Add coding rule for XML-escaping injected strings
- Fix Config.bat references in help block
- Fix stale refs and add missing functions to table
- Add missing launchers to folder structure diagram
- Add missing scopes to scope table
- Add ADR-008 for .env config file format decision
- Bump version to v2.0.0
- Update for v2.0.0
- Restore correct structure after corrupted prepend
- Update for v2.0.0
- Remove file-level header so --strip all --prepend works correctly

### Fixed
- Support single-digit M and d tokens in separator-delimited DateFormat, closes #32
- Fix month/day range validation for single-digit M and d tokens, refs #32
- Read Accounts.xlsx from Config\ instead of ReceiptsRoot
- Sanitize category names for Excel named ranges, closes #34
- XML-escape ampersand in INDIRECT SUBSTITUTE formula

### Maintenance
- Add git-cliff changelog automation, closes #30
- Add Run-SyncMonthReceipts and Run-SyncYearReceipts launchers
- Default RECEIPTS_ROOT_LOCAL to RECEIPTS_ROOT in Config.template.bat
- Add .gitattributes to enforce CRLF line endings on Windows
- Copy Accounts.xlsx to Config\ and gitignore personal data files
- Add Categories.template.json as default template
- Bump actions/checkout from 4 to 6
- Upgrade git-cliff-action from v3 to v4 to fix Debian Buster EOL build failure
- Fix git-cliff binary name in Update CHANGELOG step
- Use --prepend in git-cliff-action to avoid separate shell invocation
- Fix cliff.toml commit_parsers to use type matching, remove broken preprocessors
- Fix cliff.toml parsers to use message prefix matching instead of type field
- Fix template var commit.description -> commit.message

### Tests
- Add -DateFormat and ParseError test cases, refs #25
- Add no-method parse tests and update method-missing assertion, refs #28
- Add single-digit M-d-yy DateFormat parse tests, refs #32
- Add ConvertTo-ExcelRangeName tests, refs #34
- Assert &amp; escaping in INDIRECT formula to prevent XML parse regression

---
## [1.0.0] - 2026-03-17

### Added
- `Initialize-SyncReceipts.ps1` (via `Setup.bat`): one-time setup script that checks
  prerequisites, creates `Config\Config.bat` from the template, copies
  `Accounts.template.xlsx` to `RECEIPTS_ROOT\Accounts.xlsx`, and creates `.lnk`
  shortcuts in `RECEIPTS_ROOT`
- `Write-SyncLog` helper: centralised console logging with `[HH:mm:ss]` timestamp
  and fixed-width type tags (`STEP`, `INFO`, `WARN`, `ERROR`, `VERB`); VERB routes
  to `Write-Verbose` and is shown only with `-Verbose`
- `Get-ExcelColumnLetter` pure-PowerShell helper: converts a 1-based column index
  to an Excel column letter string (e.g. 1 -> "A", 27 -> "AA")
- Expanded Pester test suite: `Write-SyncLog`, `Get-ExcelColumnLetter`, multi-sheet
  XML patching, account-required validation for `Card`/`Checking`/`Savings`
- PSScriptAnalyzer lint tests for both `Sync-Receipts.ps1` and `Initialize-SyncReceipts.ps1`
- Comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.OUTPUTS`,
  `.EXAMPLE`) on all functions in `Sync-Receipts.ps1`
- `CONTRIBUTING.md`: prerequisites, test instructions, coding rules, commit format
- `SECURITY.md`: security policy and considerations
- `.editorconfig`: consistent indentation and line endings across file types
- Dependabot configured for weekly GitHub Actions dependency checks
- Bug report and feature request issue templates

### Changed (breaking)
- Repo reorganised into subfolders: `Scripts/`, `Launchers/`, `Tests/`, `Config/`
- `Config\Config.bat` is now the expected location for machine-local config
  (previously at repo root)
- `Config.template.bat`, `Accounts.template.xlsx`, and `Categories.json` moved
  into `Config/`
- `Parse-Receipt` renamed to `ConvertFrom-ReceiptFileName`
- `Sync-Month` renamed to `Write-MonthSheet`
- `Sync-CategorySheet` renamed to `Write-CategorySheet`
- `Setup.ps1` renamed to `Initialize-SyncReceipts.ps1`
- `Card`, `Checking`, and `Savings` payment methods now require an account number;
  filenames without one are rejected with `OK=$false`
- All `Write-Host` calls replaced with `Write-SyncLog`
- `Set-CategoryNamedRanges` builds address strings in pure PowerShell via
  `Get-ExcelColumnLetter`; no longer accesses the hidden Category sheet via COM
- Repo made public

### Fixed
- `Initialize-SyncReceipts.ps1` now correctly reads and creates `Config\Config.bat`
  (was looking at repo root after Config/ reorganisation)
- Script hang on existing workbooks: `Set-CategoryNamedRanges` previously called
  COM range methods on the hidden Category sheet, which deadlocks headless Excel
  when opening an existing workbook (closes #22)
- `Write-SyncLog` renamed from `Write-Log` to avoid `PSAvoidOverwritingBuiltInCmdlets`
  violation under `pwsh` / PowerShell Core (closes #24)
- `pushd/popd` added to all batch launchers to handle UNC path working directories
  (closes #27)

### Removed
- Deprecated Account sheet fallback from `Get-ValidAccounts`: if `Accounts.xlsx`
  is absent, account validation is now skipped (previously fell back to the Account
  sheet in the year workbook with a deprecation warning)

---

## [0.5.0] - 2026-03-17

### Added
- Per-year workbooks: each year now gets its own `{year}.xlsx` (e.g. `2026.xlsx`)
  stored directly in `RECEIPTS_ROOT`, instead of a single shared `Receipts.xlsx`
- New workbooks are created automatically on first sync for a given year

### Changed
- MAIN block refactored to group months by year; workbook open/save/close now
  happens inside a per-year loop rather than once for the entire run
- `-WorkbookPath` override still supported for testing; bypasses per-year path
- README Setup and Folder Structure updated to reflect new workbook layout
- Script version bumped to v0.5.0

### Removed
- `Receipts.xlsx` lookup fallback logic (no longer needed)

---

## [0.4.0] - 2026-03-17

### Added
- `Accounts.xlsx` support: `Get-ValidAccounts` reads account data from
  `RECEIPTS_ROOT\Accounts.xlsx` using the existing Excel COM instance
- `Accounts.template.xlsx` committed to repo with updated column schema:
  Last 4, Holder, Institution, Network, Type
- Gitignore exception so template xlsx files can be committed

### Changed
- `Get-ValidAccounts` signature updated to accept `$ReceiptsRoot` and `$Excel`;
  falls back to Account sheet in workbook with deprecation warning
- Flag message updated from "Account not in Account sheet" to "Account not in Accounts.xlsx"
- README Workbook Structure section updated to document `Accounts.xlsx` schema

---

## [0.3.0] - 2026-03-17

### Added
- `Categories.json` committed directly to repo (no template needed -- no sensitive data)
- `Get-Categories` unit tests (previously untestable without Excel COM)

### Changed
- `Get-Categories` reads from `$PSScriptRoot\Categories.json` by default;
  no longer depends on the Category sheet in `Receipts.xlsx`
- All data files renamed to CamelCase (`categories.json` -> `Categories.json`,
  `categories.template.json` -> `Categories.json`)
- `CLAUDE.md` and README updated to reflect new file locations

### Removed
- Category sheet dependency for populating dropdowns

---

## [0.2.0] - 2026-03-16

### Added
- Subcategory dropdown via post-save XML patching (`Set-SubcategoryValidationXml`):
  injects `<dataValidations>` and fixes zip binary headers
- Pester test suite: `Parse-Receipt`, `Set-SubcategoryValidationXml`, lint via PSScriptAnalyzer
- `-WorkbookPath` parameter to override the default workbook location
- VS Code tasks for running tests and syncing current/all months
- `RECEIPTS_ROOT_LOCAL` config variable for local drive equivalent of UNC path
- Test helpers: `New-MinimalXlsx`, `Get-XlsxEntry`
- XML insertion order tests (dataValidations before hyperlinks/tableParts)

### Changed
- `Run-SyncAllReceipts.bat` and docs updated for decoupled script/data layout
- `ShouldProcess` support added to all `Set-` functions
- Test batch file uses `RECEIPTS_ROOT_LOCAL` to avoid UNC path latency

### Fixed
- XML `<dataValidations>` insertion order: must precede `<hyperlinks>` and `<tableParts>`

---

## [0.1.0] - 2026-03-16

### Added
- Initial release: `Sync-Receipts.ps1` core automation
- `Parse-Receipt`: regex-parses receipt filenames into date, vendor, amount, method, account
- `Get-ValidAccounts`: reads account numbers from the Account sheet
- `Get-Categories`: reads category/subcategory data from the Category sheet
- `Set-CategoryNamedRanges`: creates named ranges for dropdown validation
- `Sync-Month`: creates/overwrites a month sheet with a 9-column formatted table
- File hyperlinks, currency formatting, date formatting
- `Run-SyncReceipts.bat` and `Run-SyncAllReceipts.bat` launchers
- `Config.template.bat` for machine-specific setup
