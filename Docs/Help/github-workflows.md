# GitHub Workflows & CI/CD

Automated checks, CI pipeline, release process, and merge requirements.

## Workflow Overview

```
Commit
  [down] (pre-commit hook)
  [|---> ASCII check
  [|---> PSScriptAnalyzer lint
  [`---> Sensitive data scan
        [down] (if pass)
        Commit succeeds
            [down]
        Push
          [down] (pre-push hook)
          [`---> Full Pester (unit + integration)
              [down] (if pass)
              github.com
                [down]
              (GitHub Actions trigger)
              [|---> commit-lint.yml (enforces Conventional Commits)
              [|---> tests.yml (runs PSScriptAnalyzer)
              [`---> labeler.yml (auto-labels PR)
                  [down] (if all pass)
                  PR mergeable
                    [down]
                  Merge
                    [down]
                  auto-release.yml (checks for [Unreleased] in CHANGELOG)
                    [|---> If found: stamps version, tags commit
                    [`---> Triggers release.yml
                        [down]
                        release.yml
                          [`---> Publishes GitHub Release
```

---

## Local Hooks (Pre-Commit & Pre-Push)

### Pre-Commit Hook

Location: `.git/hooks/pre-commit` (auto-installed by Initialize-SyncReceipts.ps1)

**Runs when:** `git commit`

**Checks:**
1. **ASCII validation** -- all .ps1 files must be pure ASCII (no smart quotes, em-dashes)
2. **PSScriptAnalyzer** -- PowerShell style & syntax lint
3. **Sensitive data scan** -- detects SSN, account numbers, credit cards, etc.

**Delegates to:** `Scripts/hooks/Invoke-PreCommitCheck.ps1`

**Exits with:**
- `0` (success) -- all checks pass
- `1` (failure) -- blocks commit, shows errors

**Override (not recommended):**
```bash
git commit --no-verify  # DO NOT USE without explicit permission
```

### Pre-Push Hook

Location: `.git/hooks/pre-push` (auto-installed by Initialize-SyncReceipts.ps1)

**Runs when:** `git push origin <branch>`

**Checks:**
- Full Pester test suite (unit + integration tests)
- PSScriptAnalyzer lint (same as pre-commit)

**Delegates to:** `Scripts/hooks/Invoke-PrePushCheck.ps1`

**Exits with:**
- `0` (success) -- all tests pass
- `1` (failure) -- blocks push

**Override (not recommended):**
```bash
git push --no-verify  # DO NOT USE without explicit permission
```

### Commit-Msg Hook

Location: `.git/hooks/commit-msg` (auto-installed by Initialize-SyncReceipts.ps1)

**Runs when:** `git commit`

**Checks:**
- Enforces Conventional Commits format: `type(scope): description`

**Exits with:**
- `0` (success) -- message format valid
- `1` (failure) -- blocks commit

**Examples:**
```
[OK] feat(categories): add filtering
[OK] fix(convert-receipt): handle slashes in vendor names
[OK] docs(readme): clarify setup steps
[X] Fixed a bug              (no type or scope)
[X] feat: add feature         (no scope)
```

---

## GitHub Actions Workflows

All workflows live in `.github/workflows/`.

### commit-lint.yml

**Trigger:** Pull request opened or updated

**Purpose:** Enforce Conventional Commits format on PR titles/commits

**Runs:** Commitlint action (npm package)

**Exits with:**
- [OK] All commits follow Conventional Commits
- [X] One or more commits violate format -> PR cannot merge

**Note:** Local `commit-msg` hook enforces same rules; this is a safety net.

---

### tests.yml

**Trigger:** Push to any branch, pull request

**Purpose:** Run Pester tests on CI environment (Windows)

**Runs on:** `windows-latest` (GitHub-hosted runner)

**Runs:**
```powershell
Invoke-Pester -Path Tests/Lint.Tests.ps1
```

**Note:** Only runs linting (PSScriptAnalyzer), NOT integration tests.
- Integration tests require Excel (not available in CI)
- Developers must run `pre-push` hook to verify integration tests locally

**Exits with:**
- [OK] No PSScriptAnalyzer violations
- [X] Linting failed -> PR cannot merge

---

### labeler.yml

**Trigger:** Pull request opened or updated

**Purpose:** Auto-apply labels to PRs based on changed files

**Configuration:** `.github/labeling.yml`

**Example:**
```yaml
- label: "scope:categories"
  files:
    - "Scripts/Sync-Receipts.ps1"
    - "Config/Categories.json"
    - "Config/Templates/Categories.template.json"
```

When PR changes any of these files -> auto-labels with `scope:categories`

**Labels available:**
- `scope:ps1`, `scope:logging`, `scope:categories`, `scope:accounts`, etc.
- `type:feature`, `type:bug`, `type:docs`, etc.
- `status:blocked`, `status:ready-to-merge`, etc.

---

### auto-release.yml

**Trigger:** Push to `master` branch (after PR merge)

**Purpose:** Watch for CHANGELOG entries and auto-tag release

**Logic:**
1. Check if CHANGELOG.md has `[Unreleased]` section with content
2. If found:
   - Extract version from changelog entry structure
   - Determine MAJOR/MINOR/PATCH bump
   - Create git tag (e.g., `v2.1.0`)
   - Push tag to remote
3. If not found:
   - Do nothing (no release)

**Triggers:** `release.yml` (when tag pushed)

**Example trigger:**
```markdown
## [Unreleased]

### Added
- New category filtering feature

### Fixed
- Account validation bug
```

After merge -> auto-release.yml detects `[Unreleased]` -> creates tag -> publishes release

---

### release.yml

**Trigger:** Git tag pushed (usually by `auto-release.yml`)

**Purpose:** Publish GitHub Release

**Logic:**
1. Extract version from git tag name (e.g., `v2.1.0`)
2. Find matching CHANGELOG.md entry
3. Use changelog entry as GitHub Release body
4. Create release on GitHub

**Example:**
- Tag: `v2.1.0`
- Finds CHANGELOG section: `## [2.1.0] - 2026-03-23`
- Creates GitHub Release with that section as body

**Manual trigger:**
```bash
git tag v2.1.0
git push origin v2.1.0
```

---

## Merge Requirements

A PR can only merge when:

1. [OK] `commit-lint` checks pass (Conventional Commits format)
2. [OK] `tests` checks pass (PSScriptAnalyzer lint)
3. [OK] No merge conflicts
4. [OK] PR review approved (if required by branch protection rules)

**Note:** Local `pre-push` hook should catch most issues before PR creation.

---

## Release Process Checklist

### Before Merging to Master

In your feature branch:

1. **Write CHANGELOG entry** at top of `CHANGELOG.md`
   ```markdown
   ## [Unreleased]

   ### Added
   - Category filtering (closes #42)

   ### Fixed
   - Account validation with empty Last4
   ```

2. **Commit CHANGELOG:**
   ```bash
   git add CHANGELOG.md
   git commit -m "docs(changelog): prepare v2.1.0 release"
   ```

3. **Push branch and open PR:**
   ```bash
   git push -u origin feat/my-feature
   gh pr create --title "feat: my feature" ...
   ```

### After Approval & Green CI

1. **Merge PR** (via `/merge-pr` or GitHub UI)
   - Triggers `auto-release.yml` on master

2. **auto-release.yml runs:**
   - Detects `[Unreleased]` in CHANGELOG
   - Creates tag `v2.1.0`
   - Pushes tag

3. **release.yml runs (on tag push):**
   - Publishes GitHub Release
   - Uses CHANGELOG entry as body

### Verify Release

```bash
gh release list  # See published releases
gh release view v2.1.0  # View release details
```

---

## Troubleshooting CI Failures

### commit-lint fails

**Error:** "Commit does not follow Conventional Commits format"

**Fix:** Ensure commit message is `type(scope): description`
```bash
git commit --amend -m "feat(categories): add filtering"
git push --force-with-lease
```

### tests fails (PSScriptAnalyzer)

**Error:** "Rule violation detected"

**Fix:** Run locally, fix issues:
```powershell
Invoke-ScriptAnalyzer -Path Scripts/Sync-Receipts.ps1
```

### pre-push hook blocks push (Pester test fails)

**Error:** "Test failed: should parse simple receipt"

**Fix:** Debug and fix test/code:
```powershell
Invoke-Pester -Path Tests/ConvertFrom-ReceiptFileName.Tests.ps1 -Verbose
# Fix code/test
git commit -m "fix(tests): edge case handling"
git push
```

---

## Disabling Hooks (Temporary)

**Caveat:** Only in emergencies; not recommended.

```bash
# Skip pre-commit and commit-msg hooks (NOT RECOMMENDED)
git commit --no-verify -m "emergency fix"

# Skip pre-push hook
git push --no-verify

# Both
git push --no-verify
```

**Re-enable hooks:**
```bash
# Hooks are auto-installed; just enable them again
Scripts/Install-GitHooks.ps1
```

---

## Related

- [workflow.md](workflow.md) -- Issue tracking, PR process, release checklist
- [commit-reference.md](commit-reference.md) -- Commit message format
- [testing-strategy.md](testing-strategy.md) -- Pester tests, CI behavior
- [coding-rules.md](coding-rules.md) -- ASCII requirements (pre-commit check)
