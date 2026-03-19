# ADR-005: Store Personal Data Files in Config/ and Gitignore Them

## Status
Accepted

## Context
`Accounts.xlsx` (personal account numbers) and `Categories.json` (personal spending
categories) are required at runtime but contain personal financial data that must
not be committed to a public repository. Originally `Accounts.xlsx` was read from
`ReceiptsRoot`, tying a personal file to the data directory. `Categories.json` was
read from `$PSScriptRoot`. Both locations were inconsistent and exposed personal data
to accidental commits. (Issue #35.)

## Decision
Both files live in `Config/` at the repo root and are listed in `.gitignore`. Committed
template files (`Accounts.template.xlsx`, `Categories.template.json`) provide the
schema and are copied to their live names during setup. The script resolves both paths
from `$PSScriptRoot` (the `Scripts/` folder), walking up one level to reach `Config/`:

    Join-Path (Split-Path $PSScriptRoot -Parent) "Config"

If a file is absent the script logs a warning and skips that validation step rather
than failing.

## Alternatives considered
- **Read from ReceiptsRoot** -- original location for `Accounts.xlsx`. Rejected
  because `ReceiptsRoot` is a data directory on a network share; placing config files
  there couples machine-specific setup to the share and risks committing personal data
  if the share is ever added to a repo.
- **Read from a user profile path (`$env:USERPROFILE`)** -- fully outside the repo.
  Rejected because it breaks when the repo is cloned to a new machine (no automatic
  template copy) and makes the config location non-obvious to new users.
- **Environment variables** -- pass account/category data via env vars set in
  `Config.bat`. Rejected because structured data (multi-row account lists, nested
  category/subcategory trees) does not map cleanly to environment variables.

## Consequences
- First-time setup must copy templates to live names (handled by `Initialize-SyncReceipts.ps1`).
- Personal data is never accidentally committed; `.gitignore` enforces this.
- `Get-ValidAccounts` no longer accepts a `-ReceiptsRoot` parameter; the path is
  fully internal, which is a breaking change from the pre-#35 interface.
- Both files are optional at runtime; the script degrades gracefully if absent.
