# Slash Commands Reference

Quick guide to all available `/command` shortcuts in Claude Code for this project.

## Available Commands

All commands are defined in `.claude/commands/` and invoked as `/command-name` in Claude Code chat.

### Workflow Commands

#### `/prep-pr` -- Prepare Pull Request

**What it does:**
1. Drafts a CHANGELOG entry (hand-crafted)
2. Commits the CHANGELOG
3. Pushes branch to remote
4. Opens a GitHub PR

**When to use:**
- Your feature branch is complete and ready for review
- CHANGELOG entry is written
- Tests pass locally

**Example:**
```
User: /prep-pr
-> Claude drafts CHANGELOG entry
-> Creates commit "docs(changelog): release v2.1.0"
-> Pushes to origin/feat/my-feature
-> Opens PR with auto-generated description
```

---

#### `/merge-pr` -- Merge Pull Request

**What it does:**
1. Checks all CI status (commit-lint, tests, labeler)
2. If all [OK], merges PR to master
3. Deletes feature branch
4. Triggers auto-release workflow

**When to use:**
- PR has been approved
- All CI checks are passing
- Ready to release

**Example:**
```
User: /merge-pr
-> Checks CI status
-> [OK] All passing
-> Merges PR
-> auto-release.yml triggers (if [Unreleased] in CHANGELOG)
```

---

### Issue & Investigation Commands

#### `/open-issues` -- List & Analyze Issues

**What it does:**
1. Fetches all open GitHub issues
2. Analyzes and categorizes them
3. Produces structured report (bugs, features, docs, etc.)
4. Identifies duplicates or related issues

**When to use:**
- Starting a new task (check if already tracked)
- Planning sprint
- Understanding backlog

**Example:**
```
User: /open-issues
-> Lists all open issues with status
-> Groups by scope (categories, accounts, tests, etc.)
-> Shows blockers and dependencies
```

---

### Review Commands

All review commands scan codebase for issues, inconsistencies, and gaps. They can fix inline and create GitHub issues for significant findings.

#### `/review-all` -- Run All Review Checks

**What it does:**
Runs all review commands in sequence:
1. `review-changelog` -- CHANGELOG accuracy
2. `review-claude` -- CLAUDE.md format and completeness
3. `review-commands` -- Slash commands (.claude/commands/)
4. `review-config` -- Config templates and schema
5. `review-docs` -- Documentation completeness
6. `review-help` -- PowerShell comment-based help
7. `review-launchers` -- Batch launcher files
8. `review-security` -- Security posture
9. `review-tests` -- Pester test coverage
10. `review-tooling` -- CI/CD and hooks

**When to use:**
- Before release (full audit)
- After major refactoring
- Quarterly maintenance

**Example:**
```
User: /review-all
-> Runs 10 review checks
-> Reports violations inline
-> Creates GitHub issues for significant findings
```

---

#### `/review-changelog` -- Audit CHANGELOG.md

**What it does:**
1. Validates CHANGELOG format (Keep a Changelog)
2. Checks version numbers (Semantic Versioning)
3. Verifies section structure (Added, Changed, Fixed, Removed)
4. Looks for completeness (all commits documented?)
5. Checks for duplicates or inconsistencies

**Fixes:** Minor formatting issues inline

**When to use:**
- After writing CHANGELOG entry
- Before tagging release
- Quarterly audit

---

#### `/review-claude` -- Audit CLAUDE.md Files

**What it does:**
1. Validates project CLAUDE.md structure
2. Checks token efficiency (unnecessary duplication?)
3. Verifies all required sections present
4. Checks links (do they point to valid files?)
5. Looks for outdated information

**Fixes:** Reformat for efficiency; remove duplicates

**When to use:**
- After major architectural change
- If CLAUDE.md grows beyond ~200 lines
- Quarterly audit

**Note:** This is why we extracted separate `.md` files (token optimization).

---

#### `/review-commands` -- Audit Slash Commands

**What it does:**
1. Lists all commands in `.claude/commands/`
2. Checks each command for stale or missing entries
3. Verifies descriptions match current behavior
4. Looks for unimplemented commands
5. Checks for duplicates

**Fixes:** Updates descriptions; removes stale commands

**When to use:**
- After adding new command
- If workflow changes
- Quarterly audit

---

#### `/review-config` -- Audit Configuration

**What it does:**
1. Validates Config.template.ini schema
2. Checks Accounts.template.xlsx structure
3. Validates Categories.template.json format
4. Checks Methods.template.json syntax
5. Validates SensitivePatterns.template.json regex
6. Verifies all templates are in sync with code

**Fixes:** Schema corrections; regex fixes

**When to use:**
- After schema change
- Before committing template updates
- Quarterly audit

---

#### `/review-docs` -- Audit Documentation

**What it does:**
1. Checks Docs/ and ADRs/ for completeness
2. Verifies all features are documented
3. Looks for broken links
4. Checks for outdated information
5. Verifies code examples are current

**Fixes:** Updates links; removes stale content

**When to use:**
- After feature addition
- When documentation feels incomplete
- Quarterly audit

---

#### `/review-help` -- Audit PowerShell Help

**What it does:**
1. Scans all PowerShell files for comment-based help
2. Checks .SYNOPSIS, .DESCRIPTION, .PARAMETER sections
3. Verifies parameter names match function signature
4. Looks for missing examples
5. Checks help is current with function behavior

**Fixes:** Adds missing help blocks; updates examples

**When to use:**
- After adding new PowerShell function
- If function signature changes
- Quarterly audit

---

#### `/review-launchers` -- Audit Batch Files

**What it does:**
1. Scans all `.bat` files in Launchers/ and Tests/
2. Checks for correct PowerShell invocation
3. Verifies paths are correct
4. Looks for consistent error handling
5. Checks for documentation in comments

**Fixes:** Corrects paths; adds comments

**When to use:**
- After adding new launcher
- If folder structure changes
- Quarterly audit

---

#### `/review-security` -- Audit Security Posture

**What it does:**
1. Scans code for secrets (passwords, API keys, etc.)
2. Checks SECURITY.md for completeness
3. Verifies SensitivePatterns.json covers common PII
4. Looks for insecure patterns (plaintext passwords, etc.)
5. Checks pre-commit hook is enforcing data protection

**Fixes:** Removes found secrets; suggests patterns

**Issues:** Creates GitHub issue for significant findings

**When to use:**
- Before release
- Quarterly security audit
- If handling new types of sensitive data

---

#### `/review-tests` -- Audit Pester Tests

**What it does:**
1. Lists all test files in Tests/
2. Checks which functions have unit tests ([OK] or [X])
3. Verifies test file naming convention
4. Checks test coverage (pure functions tested?)
5. Looks for stale or skipped tests
6. Verifies fixture files exist

**Fixes:** Updates test documentation; removes skipped tests

**When to use:**
- After adding new function
- When test coverage feels incomplete
- Quarterly audit

---

#### `/review-tooling` -- Audit CI & Hooks

**What it does:**
1. Validates GitHub Actions workflows (tests.yml, commit-lint.yml, etc.)
2. Checks pre-commit and pre-push hook scripts
3. Verifies hooks are installed correctly
4. Looks for missing workflow conditions
5. Checks for deprecated GitHub Actions syntax

**Fixes:** Updates workflow syntax; fixes hook logic

**When to use:**
- After updating CI/CD
- If GitHub Actions API changes
- Quarterly audit

---

## Command Usage Tips

### Before Starting Work

```bash
/open-issues            # See what's tracked
# Pick an issue or create one for your task
```

### During Development

```bash
/review-tests           # Make sure your new function has tests
/review-docs            # Update docs as you code
```

### Before Commit

```bash
git status              # Check what's staged
# Write commit message
# Rely on pre-commit hook to catch issues
```

### Before PR

```bash
/prep-pr                # Draft CHANGELOG, commit, push, open PR
```

### After Approval

```bash
/merge-pr               # Check CI, merge, release
```

### Before Release

```bash
/review-all             # Full audit before tagging
```

---

## Command Output

Most review commands output:

```
[OK] Found no issues
or
[X] Found N issues:
  [issue 1]
  [issue 2]
  ...
  [link] Created GitHub issue #123 for: [significant finding]
```

---

## Creating New Commands

To add a new command:

1. Create `.claude/commands/my-command.md`
2. Write task description (reusable prompt)
3. Reference in CLAUDE.md slash commands table
4. Test with `/my-command` in chat

Example:
```markdown
# My Custom Command

Reusable task: [description]

## Steps
1. Step 1
2. Step 2
3. Step 3

## Output
[Expected output format]
```

---

## Related

- [workflow.md](workflow.md) -- `/prep-pr` and `/merge-pr` in context
- [../context-selection.md](../context-selection.md) -- When to use each review command
- `.claude/commands/` -- Source of all available commands
