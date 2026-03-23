# Workflow & Issue Tracking

Guide for issue creation, PR process, and release workflow.

## When to Create an Issue

**Create issues automatically for non-trivial changes:**
- New features
- Bug fixes
- Refactoring with side effects
- Documentation gaps

**Skip issues only for:**
- Single-word typo fixes
- Trivial one-line logic adjustments

## Before Starting Work

1. **Scan open issues** to avoid duplicating tracked work
2. **Identify affected issues** -- your change may close or relate to existing issues
3. **Note the issue number** for the commit message

## Making a Commit

When an issue exists, link it in the commit message body:

```
feat(categories): add category filtering

This implements dynamic category selection as requested.

Closes #42
```

Or for related (non-closing) changes:
```
Refs #25
```

## Issue Resolution

When closing an issue:

1. **Add a resolution comment** describing:
   - What was done
   - Which files changed
   - Why that approach was chosen

   ```bash
   gh issue comment <N> --body "<<'EOF'
   ## Resolution

   Implemented category filtering in Get-Categories function.

   **Files changed:**
   - Scripts/Sync-Receipts.ps1 (added -Filter parameter)
   - Tests/Get-Categories.Tests.ps1 (added 3 test cases)

   **Approach:** Used Where-Object for runtime filtering instead of
   loading config differently, allowing flexibility without refactoring
   the config structure.
   EOF
   "
   ```

2. **Close with reason:**
   ```bash
   gh issue close <N> --reason completed
   ```
   Or for won't-fix:
   ```bash
   gh issue close <N> --reason "not planned"
   ```

## Pull Request Workflow

### 1. Prepare Branch

Use `/prep-pr` to:
- Draft a CHANGELOG entry
- Commit it
- Push branch
- Open PR

Or manually:
```bash
git checkout -b feat/my-feature
# Make changes, commit, push
gh pr create --title "..." --body "..."
```

### 2. PR Description Template

See `.github/PULL_REQUEST_TEMPLATE.md`. Include:
- What problem does this solve?
- How was it tested?
- Any breaking changes?

### 3. Check CI Status

CI must pass before merge:
- `commit-lint.yml` -- enforces Conventional Commits format
- `tests.yml` -- runs full Pester suite on `windows-latest`
- All status checks must be green

### 4. Merge PR

Use `/merge-pr` to:
- Check CI status
- Merge (if all checks pass)

Or manually:
```bash
gh pr merge --squash  # or --rebase, depending on preference
```

**After merge:** The `auto-release.yml` workflow watches for CHANGELOG entries and will automatically:
- Bump version (based on entry type: MAJOR, MINOR, PATCH)
- Tag the commit
- Trigger `release.yml` to publish the GitHub Release

## Release Workflow

### 1. Hand-Craft CHANGELOG Entry

Write at the **top** of `CHANGELOG.md` **before tagging**:

```markdown
## [Unreleased]

### Added
- New category filtering feature in Get-Categories

### Fixed
- Account validation bug when Last4 is empty

### Changed (breaking)
- Renamed -ReceiptsPath parameter to -ReceiptsRoot
```

### 2. Merge to Main

When PR merges, `auto-release.yml` detects `[Unreleased]` section and:
1. Extracts the version from CHANGELOG
2. Creates a git tag (e.g., `v2.1.0`)
3. Commits the CHANGELOG change

### 3. Release Creation

Push the tag (or let the workflow do it):
```bash
git push origin v2.1.0
```

`release.yml` automatically publishes the GitHub Release using the CHANGELOG entry as the body.

## Key Rules

- **Never edit CHANGELOG manually** during development -- write real entries at the top when releasing
- **One commit per `type(scope)`** -- no batching unrelated changes
- **Commit as work progresses** -- not all at the end
- **Propose commit message** and wait for approval before committing
- **Stage specific files** (`git add file1 file2`) not `git add -A`
- **Never bypass pre-commit hooks** with `--no-verify`

## Related

- [commit-reference.md](commit-reference.md) -- Commit format and types
- [commit-scopes.md](commit-scopes.md) -- Scope selection table
- [github-workflows.md](github-workflows.md) -- CI/CD pipeline details
- [config-schema.md](config-schema.md) -- Config initialization during setup
