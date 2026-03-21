# Changelog

All notable changes to this project will be documented in this file.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
This project uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

---

## [4.0.2] - 2026-03-21

### Added
- `Get-ReceiptFlag` pure helper extracted from `Write-MonthSheet`; evaluated parsed receipt rows against configured rules and returns a flag string. Unit-tested in `Tests/Get-ReceiptFlag.Tests.ps1`.
- `Copy-ConfigTemplate` pure helper extracted from `Initialize-SyncReceipts.ps1`; copies template files without overwriting existing ones. Unit-tested in `Tests/Initialize-SyncReceipts.Tests.ps1`.
- Integration test fixture and runner scripts in `Tests/Integration/` covering flag scenarios, accounts template schema, and setup idempotency.
- Lint.Tests.ps1 extended to cover all `Scripts/*.ps1` files (previously only covered `Sync-Receipts.ps1`).

### Fixed
- Pre-push hook now auto-installs PSScriptAnalyzer when missing, so lint tests run on first push without manual setup.
- Pre-push hook now fails correctly when a Pester test file fails to load (container/discovery error), preventing lint from being silently skipped.

---

## [4.0.1] - 2026-03-21

### Fixed
- Launcher shortcuts are now created in `WORKBOOKS_ROOT` instead of `RECEIPTS_ROOT`.
  For users with a separate workbooks location (ADR-006), shortcuts now appear alongside
  the year workbooks where they are most useful. Single-location setups are unaffected
  (`WORKBOOKS_ROOT` defaults to `RECEIPTS_ROOT` when not set).

---

## [4.0.0] - 2026-03-21

### Changed (breaking)
- Template files moved from `Config/` to `Config/Templates/`: `Accounts.template.xlsx`,
  `Categories.template.json`, `Methods.template.json`, `Config.template.ini`
  (was `Config.template.env`). Update any scripts that reference template paths directly.
  `Initialize-SyncReceipts.ps1` detects and moves leftover pre-v4.0.0 template files
  automatically on next run.
- `Config/Config.env` renamed to `Config/Config.ini`. Existing users must rename their
  local file before running the launchers. `Initialize-SyncReceipts.ps1` detects the old
  name on upgrade and prompts to rename it automatically.

---

## [3.1.0] - 2026-03-21

### Added

- `-ConfigRoot` parameter on `Sync-Receipts.ps1`: override the config directory (`Accounts.xlsx`, `Categories.json`, `Methods.json`) without touching the real `Config/` folder. Defaults to the repo's `Config/` subfolder --existing behaviour unchanged.
- `-OutputPath` parameter on `New-AccountsTemplate.ps1`: write the generated accounts template to any path instead of the fixed `Config\Accounts.template.xlsx`. Defaults to the original path --existing behaviour unchanged.

---

## [3.0.2] - 2026-03-21

### Fixed

- Corrected `Install-Module` splatting in `Initialize-SyncReceipts.ps1` so PSScriptAnalyzer installs without error on first setup
- Guarded `NumberFormat` assignment against null `DataBodyRange` when a month sheet has zero receipts
- Corrected inactive account flag cell value to `"Account inactive"` to match documented behaviour
- Amount column now uses accounting number format; Category and Subcategory column widths corrected to measured values
- `Test: Unit only` VS Code task filter corrected to exclude `Lint.Tests.ps1` properly
- Escaped regex special characters and removed XPath string interpolation to prevent injection vulnerabilities in receipt parsing
- Moved PR title into an environment variable in the `commit-lint` workflow to prevent shell injection
- Updated pre-push hook to use the Pester 5 `New-PesterConfiguration` API (was using the deprecated v4 call)
- Pinned all GitHub Actions to commit SHAs; updated `actions/checkout` to v4

---

## [3.0.1] - 2026-03-20

### Fixed

- Pre-commit hook now validates ASCII encoding in staged `.json` and `.md` files (previously `.ps1` only)
- `commit-lint` CI workflow now accepts breaking-change commits (`feat!:`, `fix!:` forms)
- `commit-lint` CI workflow now skips `fixup!` and `squash!` commits (matches local hook behaviour)
- `.gitignore` path for `Tests/sync-output.txt` was lowercase (`tests/`); fixed for case-sensitive CI
- `.gitattributes` missing line-ending rules for `*.psd1`, `*.editorconfig`, and `*.lnk`

---

## [3.0.0] - 2026-03-20

### Added
- `Get-Methods` function: reads configurable payment method tokens from
  `Config\Methods.json`; falls back to built-in defaults (`Card`, `Check`,
  `Checking`, `Savings`, `Transfer`, `Wire`) if the file is absent or invalid.
- `Config\Methods.template.json`: default method token list, committed to the
  repo. Copy to `Config\Methods.json` to customise.
- `Scripts\New-AccountsTemplate.ps1`: regenerates `Config\Accounts.template.xlsx`
  with the correct column schema and dropdown validation. Re-run whenever the
  Accounts schema changes.
- `Account inactive` flag in the Flag column when a receipt's account is present
  in `Accounts.xlsx` but has `Status = Inactive`.
- `Unrecognised method` flag in the Flag column when a receipt's method token is
  not in the loaded methods list.

### Changed (breaking)
- `Accounts.xlsx` schema restructured (see ADR-009): columns are now
  Last 4 / Method / Holder / Institution / Account / Status (was
  Last 4 / Holder / Institution / Network / Type). Rebuild `Config\Accounts.xlsx`
  from `Config\Accounts.template.xlsx`.
- `ConvertFrom-ReceiptFileName` now validates the method token against the loaded
  methods list. Unrecognised tokens produce `OK=$true` with the flag
  `Unrecognised method` rather than being silently accepted.

---

## [2.0.1] - 2026-03-20

### Added
- `Setup.bat` now automatically installs local git hooks (pre-commit, pre-push,
  commit-msg) during setup. Previously required running the installer manually.

### Fixed
- Setup script now correctly copies `Config\Categories.template.json` to
  `Config\Categories.json` on first run (step was missing in v2.0.0).

---

## [2.0.0] - 2026-03-19

### Changed (breaking)
- Config file renamed from `Config.bat` to `Config.env`. The new format is plain
  `KEY=value` text with `#` comments -- no batch syntax. To upgrade: copy your
  values from `Config\Config.bat` into a new `Config\Config.env`. Fresh installs
  via `Setup.bat` create it automatically.

### Added
- Stale subcategories are cleared on re-sync: if a row's saved Subcategory is no
  longer valid for its current Category, it is cleared automatically rather than
  silently preserved.

### Fixed
- Category names containing special characters (`&`, `/`, spaces) now work
  correctly as Excel named ranges and in the dependent subcategory dropdown.
- The subcategory INDIRECT formula now correctly XML-escapes `&` as `&amp;`,
  preventing an Excel parse error on open.

---

## [1.3.2] - 2026-03-19

### Added
- `Config\Categories.template.json` committed to the repo as the default categories
  template. Copy to `Config\Categories.json` to customise.

### Changed
- `Accounts.xlsx` moved from `RECEIPTS_ROOT` to `Config\`. Move your existing file or
  copy `Config\Accounts.template.xlsx` to `Config\Accounts.xlsx` and re-enter your
  accounts.

### Fixed
- `Accounts.xlsx` was being read from `ReceiptsRoot` instead of `Config\` after the
  v1.0.0 folder reorganisation.

---

## [1.3.1] - 2026-03-18

### Fixed
- Month and day out-of-range validation now correctly flags single-digit `M` and `d`
  tokens in separator-delimited `-DateFormat` strings (e.g. `M-d-yy`).

---

## [1.3.0] - 2026-03-18

### Added
- `-WorkbooksRoot` parameter: store per-year workbooks in a separate location from
  receipts (e.g. a local drive while receipts live on a network share). Set
  `WORKBOOKS_ROOT` in `Config\Config.bat`. `Setup.bat` now prompts for this during
  setup, defaulting to `RECEIPTS_ROOT`.

### Fixed
- `-DateFormat` separator-delimited formats (e.g. `M-d-yy`) now correctly accept
  single-digit month and day values.

---

## [1.2.0] - 2026-03-18

### Added
- `Method` and `Account` fields in receipt filenames are now optional. Rows with no
  Method are written with both fields blank and flagged `Method missing` in the Flag
  column.
- `Run-SyncMonthReceipts.bat` and `Run-SyncYearReceipts.bat` launchers: sync a
  specific month (prompts for YYMM) or all months in a specific year (prompts for
  YYYY).

---

## [1.1.0] - 2026-03-18

### Added
- `-DateFormat` parameter: configurable `.NET ParseExact` format string for the date
  portion of receipt filenames (default: `yyMMdd`). Set `DATE_FORMAT` in
  `Config\Config.bat`. Per-field parse error flags (`Could not parse filename`,
  `Month out of range`, `Day out of range`, `Invalid date`) are written to the Flag
  column when a filename date cannot be parsed.

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
