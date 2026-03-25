# Commit Scopes Reference

Scopes organize commits by subsystem and file location. Use the narrowest scope that applies.

| Scope | Files / area | Example |
|-------|-------------|---------|
| `ps1` | Sync-Receipts.ps1 (general; use a narrower scope when one applies) | `fix(ps1): handle missing file edge case` |
| `logging` | Write-SyncLog function and console output | `refactor(logging): add DEBUG tag support` |
| `categories` | Categories feature: Get-Categories, Write-CategorySheet, Config/Categories.json, Config/Templates/Categories.template.json | `feat(categories): add category filtering` |
| `accounts` | Accounts feature: Get-ValidAccounts, Config/Accounts.xlsx, Config/Templates/Accounts.template.xlsx, Scripts/New-AccountsTemplate.ps1 | `fix(accounts): validate account status field` |
| `methods` | Configurable method tokens: Get-Methods, Config/Methods.json, Config/Templates/Methods.template.json | `feat(methods): support crypto method tokens` |
| `write-month` | Write-MonthSheet function | `perf(write-month): optimize cell formatting` |
| `convert-receipt` | ConvertFrom-ReceiptFileName function | `fix(convert-receipt): handle vendor name with slashes` |
| `xml` | Set-SubcategoryValidationXml XML patching | `fix(xml): escape special characters in validation` |
| `tests` | Tests/ (per-function test files and Lint.Tests.ps1) | `test(tests): add edge case for empty input` |
| `adr` | Docs/ADRs/ | `docs(adr): add ADR-013 for caching strategy` |
| `readme` | README.md | `docs(readme): clarify setup prerequisites` |
| `claude.md` | CLAUDE.md | `docs(claude.md): reorganize context structure` |
| `ci` | GitHub Actions workflows (.github/), hook scripts (Scripts/hooks/) | `ci(ci): add build job to GitHub Actions` |
| `config` | Config.ini, Config/ subfolder, .vscode/, batch launcher files | `chore(config): add new VSCode extension` |
| `sensitive-patterns` | Config/SensitivePatterns.json, Config/Templates/SensitivePatterns.template.json, Scripts/hooks/Invoke-SensitiveDataCheck.ps1 | `feat(sensitive-patterns): detect SSN patterns` |
| `changelog` | CHANGELOG.md | `chore(changelog): release v2.0.0` |
| `contributing` | CONTRIBUTING.md | `docs(contributing): add testing guide` |
| `security` | SECURITY.md | `docs(security): update responsible disclosure email` |

## Scope Selection Examples

- **General PowerShell logic** -> use narrower scope (e.g., `convert-receipt`, `write-month`) or `ps1` as fallback
- **Configuration file** -> use corresponding scope (`categories`, `accounts`, `methods`, etc.)
- **Hook or workflow** -> use `ci`
- **Documentation** -> use corresponding scope (`readme`, `adr`, `changelog`, etc.)
- **Multiple subsystems** -> split into separate commits with appropriate scopes
