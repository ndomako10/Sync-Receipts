# Changelog

All notable changes to this project will be documented in this file.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
This project uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

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
- `config.template.bat` for machine-specific setup
