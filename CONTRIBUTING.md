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
| `config` | Config.env, Config/ subfolder, .vscode/, batch launcher files |
| `changelog` | CHANGELOG.md |
| `contributing` | CONTRIBUTING.md |
| `security` | SECURITY.md |

## Releasing

Changelog generation and GitHub Release creation are fully automated via [git-cliff](https://git-cliff.org/) and GitHub Actions. To cut a release:

1. Ensure all changes are committed and pushed to `master`
2. Bump the version in the `Scripts/Sync-Receipts.ps1` header and commit:
   ```
   docs(ps1): bump version to vX.Y.Z
   ```
3. Tag the release and push the tag:
   ```bash
   git tag vX.Y.Z
   git push origin vX.Y.Z
   ```
4. The release workflow (`.github/workflows/release.yml`) automatically:
   - Generates the changelog entry from commits since the last tag using `cliff.toml`
   - Prepends the entry to `CHANGELOG.md` and commits it to `master`
   - Creates a GitHub Release with the generated release notes

To preview what the next release entry will look like before tagging, run locally:
```bash
git cliff --unreleased
```

## Submitting a Pull Request

1. Fork the repo and create a branch from `master`
2. Make your changes and ensure `Invoke-Pester Tests/ -Output Detailed` passes
3. Follow the commit message format above
4. Open a pull request against `master`

## Reporting Issues

Use the issue templates on GitHub -- bug reports and feature requests each have a dedicated template.
