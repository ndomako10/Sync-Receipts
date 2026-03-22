Review the codebase for security issues, sensitive data exposure, and insecure patterns.

Read each of the following files in full:
- Scripts/Initialize-SyncReceipts.ps1
- Scripts/Install-GitHooks.ps1
- Scripts/New-AccountsTemplate.ps1
- Scripts/hooks/Invoke-PreCommitCheck.ps1
- Scripts/hooks/Invoke-PrePushCheck.ps1
- Scripts/hooks/Invoke-CommitMsgCheck.ps1
- .github/workflows/tests.yml
- .github/workflows/release.yml
- .github/workflows/commit-lint.yml
- .gitignore
- Config/Templates/Config.template.ini

Run these targeted scans (covers Scripts/Sync-Receipts.ps1 without a full read):
- `git log --oneline -20`
- `git grep -i "password\|secret\|token\|api.key\|credential" -- "*.ps1" "*.yml" "*.json" "*.md"`
- `git grep -iE "(Card|Checking|Savings|Check|Wire|Transfer)\s+[0-9]{4}" -- "*.md" "*.ps1" "*.json"`
- `git grep -n "Invoke-Expression\|iex \|& \$\|\[scriptblock\]::Create" -- "*.ps1"` to scan for command injection patterns
- `git grep -n "param\b.*Path\|ReceiptsRoot\|WorkbooksRoot\|RootPath" Scripts/Sync-Receipts.ps1` to identify where external paths are accepted
- `git grep -n "Resolve-Path\|Test-Path\|Join-Path" Scripts/Sync-Receipts.ps1 | head -30` to check path validation patterns

Check each area and flag any issues:

1. **Sensitive data in committed files** -- Scan for real account numbers, institution names, personal names, email addresses, or credentials that should not be in the repo. Flag anything that looks like real personal data rather than a placeholder (e.g. `1234`, `xxxx`, `MyBank`).

2. **Gitignore coverage** -- Are all personal data files gitignored?
   - `Config/Config.ini`
   - `Config/Accounts.xlsx`
   - `Config/Categories.json`
   - `Config/Methods.json`
   - Any output files or logs that could contain receipt data
   Verify each is present in `.gitignore`.

3. **Credential patterns in code** -- Check scripts and workflows for hardcoded secrets, tokens, or credentials. Flag any `$env:` variables that should be secrets but are not using GitHub Actions secrets.

4. **Command injection** -- Check any place where external input (filenames, config values, user-provided paths) is passed to shell commands, `Invoke-Expression`, or string interpolation in a way that could allow injection.

5. **File path traversal** -- Check scripts that accept path parameters for validation. Are paths resolved and checked before use? Could a malformed path escape the intended directory?

6. **Hook security** -- Can the pre-commit or pre-push hooks be trivially bypassed (e.g. by staging a file that disables the hook)? Are there any patterns in `Invoke-PreCommitCheck.ps1` that could be exploited by a malicious file in the working tree?

7. **CI/CD security** -- In GitHub Actions workflows:
   - Are third-party actions pinned to a commit SHA or at minimum a version tag?
   - Are secrets accessed only via `${{ secrets.* }}` and never echoed to logs?
   - Are pull request triggers restricted appropriately (e.g. `pull_request` vs `pull_request_target`)?

8. **Open issues** -- Run `gh issue list --state open --label security` and check whether any open security issues are already addressed or remain outstanding.

## Output format

For each finding: `SEVERITY | FILE:LINE | one-sentence problem | one-sentence fix`
Severity: CRITICAL (broken/data-loss), MAJOR (correctness gap), MINOR (polish).
If no issues in an area: `OK: <area>`
No narrative. No file contents. Findings only.
