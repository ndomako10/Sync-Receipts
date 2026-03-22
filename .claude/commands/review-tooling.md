Review all tooling, automation, and CI/CD configuration for correctness and consistency.

Read each of the following files in full:
- .github/workflows/tests.yml
- .github/workflows/release.yml
- .github/workflows/commit-lint.yml
- .github/workflows/labeler.yml
- .github/dependabot.yml
- .github/labeling.yml
- Scripts/hooks/pre-commit
- Scripts/hooks/pre-push
- Scripts/hooks/commit-msg
- Scripts/hooks/Invoke-PreCommitCheck.ps1
- Scripts/hooks/Invoke-PrePushCheck.ps1
- Scripts/hooks/Invoke-CommitMsgCheck.ps1
- Scripts/Install-GitHooks.ps1
- Scripts/Initialize-SyncReceipts.ps1 (steps 7-8 only: module install and hook install)
- .vscode/tasks.json
- .gitattributes
- CONTRIBUTING.md (Local Git Hooks section only)

Also run:
- `git log --oneline -10` to see recent commits
- `ls .github/ISSUE_TEMPLATE/` to list issue templates
- `gh issue list --state open --label ci` to check for known outstanding tooling gaps

Check each area against the actual repo state and flag any issues:

1. **CI workflows** --" For each workflow in `.github/workflows/`:
   - Do all file paths referenced (`Scripts/`, `Tests/`, `Config/`) exist in the repo?
   - Are `actions/checkout` and any other actions pinned to reasonable versions?
   - Does `tests.yml` use `New-PesterConfiguration` (required for Pester 5 when combining PassThru + CodeCoverage)?
   - Does `release.yml` extract version via `$GITHUB_OUTPUT` and guard empty grep results with `-z`?
   - Does `commit-lint.yml` duplicate validation that is already in the `commit-msg` local hook? Note overlap but do not flag as an error --" belt-and-suspenders is intentional.

2. **Local git hooks** --" Cross-reference three sources: `Scripts/hooks/` files on disk, the install loop in `Scripts/Install-GitHooks.ps1`, and the hooks table in `CONTRIBUTING.md`:
   - Every hook file in `Scripts/hooks/` (shell entry points only: `pre-commit`, `pre-push`, `commit-msg`) must appear in the `Install-GitHooks.ps1` install loop.
   - Every hook in the CONTRIBUTING.md table must have a corresponding file in `Scripts/hooks/`.
   - Hook shell scripts must delegate to a PS1 file via `powershell` or `pwsh`; flag any shell logic beyond finding the executable and calling the PS1.

3. **`.gitattributes`** --" Every shell hook file in `Scripts/hooks/` (files with no extension: `pre-commit`, `pre-push`, `commit-msg`) must have a `text eol=lf` rule. Flag any hook file that is missing one.

4. **VS Code tasks** --" For each task in `.vscode/tasks.json`:
   - Do all hardcoded file paths (Scripts/, Tests/, Config/, Launchers/) exist?
   - Does the "Lint: PSScriptAnalyzer" task have an inline `problemMatcher` with `fileLocation: ["absolute"]`?
   - Does the "Release: Bump version" task exist? Note: per CLAUDE.md, version is tracked in git tags only and the script header carries no version number --" flag the task as stale if it attempts to patch a version string into the script.
   - Does the "Release: Tag and push" task guard against tagging when `git tag` fails?

5. **`dependabot.yml`** --" Does it cover the `github-actions` ecosystem? Are update intervals reasonable?

6. **`Scripts/Initialize-SyncReceipts.ps1`** --" Does the module install loop (step 7) install both `Pester` (MinVersion 5.0) and `PSScriptAnalyzer`? Does step 8 call `Install-GitHooks.ps1`?

7. **Stale references** --" Any workflow, task, or hook file referencing old paths (e.g. `Config.bat`, `Config.template.bat`, old script names) or parameters that no longer exist?

8. **Missing automation** --" Based on the current repo, flag any obvious gaps:
   - Is there a labeler workflow and does `labeling.yml` cover all key file patterns (Scripts/, Tests/, Docs/, Config/, .github/)?
   - Do the ISSUE_TEMPLATE files exist and appear complete?

9. **Open issues** --" For each issue returned by `gh issue list --state open --label ci`, state whether it is already addressed by the current tooling configuration (flag for closing if so) or remains outstanding.

## Output format

For each finding: `SEVERITY | FILE:LINE | one-sentence problem | one-sentence fix`
Severity: CRITICAL (broken/data-loss), MAJOR (correctness gap), MINOR (polish).
If no issues in an area: `OK: <area>`
No narrative. No file contents. Findings only.
