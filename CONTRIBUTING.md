# Contributing

## Prerequisites

- Windows 10/11
- PowerShell 5.0+
- Microsoft Excel 2016+ (required to run the script; not required to run tests)
- [Pester 5+](https://pester.dev) and [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer) (for tests)

Install the test dependencies:

```powershell
Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser
Install-Module PSScriptAnalyzer -Scope CurrentUser
```

## Running Tests

```powershell
Invoke-Pester Tests/ -Output Detailed
```

All tests run without Excel. Excel COM integration tests are local-only and not part of the automated suite.

## Coding Rules

- **Error handling** -- every new code block needs `try/catch` and `Write-SyncLog` calls.
  Use `-Tag WARN` for non-fatal warnings, `-Tag ERROR` for errors, `-Tag VERB` for diagnostic detail.
- **No `$variable:` in strings** -- PowerShell interprets the colon as a drive separator.
  Use `${variable}:` instead.
- **ASCII only** -- the script runs from a network share; non-ASCII characters break parsing.
  Verify after editing: `[System.Text.Encoding]::ASCII.GetByteCount($content) -eq $content.Length`

## Commit Messages

This project uses [Conventional Commits](https://www.conventionalcommits.org):

```
type(scope): short description
```

| Type | When |
|------|------|
| `feat` | New feature |
| `fix` | Bug fix |
| `docs` | Documentation only |
| `test` | Test additions or changes |
| `refactor` | Code change with no behaviour change |
| `chore` | Tooling, gitignore, config |

| Scope | Area |
|-------|------|
| `ps1` | Sync-Receipts.ps1 (general) |
| `logging` | Write-SyncLog and console output |
| `categories` | Categories feature: Get-Categories, Write-CategorySheet, Config/Categories.json, Config/Categories.template.json |
| `accounts` | Accounts feature: Get-ValidAccounts, Config/Accounts.xlsx, Config/Accounts.template.xlsx |
| `write-month` | Write-MonthSheet function |
| `convert-receipt` | ConvertFrom-ReceiptFileName function |
| `xml` | Set-SubcategoryValidationXml |
| `tests` | Tests/ |
| `readme` | README.md |
| `ci` | GitHub Actions |
| `claude.md` | CLAUDE.md |
| `changelog` | CHANGELOG.md |
| `security` | SECURITY.md |
| `config` | Config.env, Config/ subfolder, .vscode/, batch launcher files |
| `contributing` | CONTRIBUTING.md |

## Releasing

GitHub Release creation is automated via GitHub Actions. To cut a release:

1. Write the changelog entry at the top of `CHANGELOG.md`, before all existing entries.
   Follow the [Keep a Changelog](https://keepachangelog.com) format -- user-facing language,
   grouped under `### Added`, `### Changed (breaking)`, `### Fixed`, `### Removed` as needed.
   End the entry with a `---` separator line.
2. Bump the version in the `Scripts/Sync-Receipts.ps1` header (line 1 and `.NOTES`) and
   commit both changes together:
   ```
   docs(ps1): bump version to vX.Y.Z
   ```
3. Ensure all changes are committed and pushed to `master`.
4. Tag the release and push the tag:
   ```bash
   git tag vX.Y.Z
   git push origin vX.Y.Z
   ```
5. The release workflow (`.github/workflows/release.yml`) automatically reads the top
   section of `CHANGELOG.md` (everything before the first `---`) and creates a GitHub
   Release with that text as the release notes.

## Submitting a Pull Request

1. Fork the repo and create a branch from `master`
2. Make your changes and ensure `Invoke-Pester Tests/ -Output Detailed` passes
3. Follow the commit message format above
4. Open a pull request against `master`

## Reporting Issues

Use the issue templates on GitHub -- bug reports and feature requests each have a dedicated template.
