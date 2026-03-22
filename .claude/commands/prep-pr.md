# /prep-pr

Prepare and open a pull request for the current branch. Drafts a CHANGELOG entry if the
changes are user-facing, commits it, pushes the branch, and opens the PR.

Run this command when your branch is ready for review.

## Step 1 -- Verify branch

Confirm the current branch is not master:
```bash
git branch --show-current
```
If on master, stop and tell the user to run this on a feature branch.

## Step 2 -- Read commits on this branch

```bash
git log master..HEAD --format="%s"
```

## Step 3 -- Determine if changes are user-facing

User-facing commit types: `feat`, `fix`, `perf`
Non-user-facing: `ci`, `test`, `chore`, `refactor`, `style`, `build`, `docs`

If ALL commits are non-user-facing, skip to Step 6 (no CHANGELOG entry needed).
If ANY commit is user-facing, proceed to Step 4.

## Step 4 -- Draft CHANGELOG entry

Group the user-facing commits into Keep a Changelog sections:
- `feat` -> `### Added`
- `fix` -> `### Fixed`
- `perf` -> `### Changed`

Write in user-facing language -- not commit subjects verbatim. Omit empty sections.

Draft format:
```
## [Unreleased]

### Added
- ...

### Fixed
- ...

---
```

Show the draft to the user and ask for approval or edits before writing it.

## Step 5 -- Write and commit CHANGELOG

Insert the approved entry at the top of CHANGELOG.md (below the header, replacing any
existing `## [Unreleased]` section if present).

Stage and commit:
```bash
git add CHANGELOG.md
git commit -m "docs(changelog): add [Unreleased] entry"
```

## Step 6 -- Scan open issues

```bash
gh issue list --state open --json number,title,labels --limit 100
```

Identify any issues this PR closes or relates to. Use `Closes #N` for issues resolved
by this PR, `Refs #N` for related ones.

## Step 7 -- Push branch

```bash
git push -u origin HEAD
```

## Step 8 -- Open PR

Use the PR title from the branch's primary commit type and description. Follow
Conventional Commits format (e.g. `feat(ps1): add multi-currency support`).

```bash
gh pr create --title "<title>" --body "$(cat <<'EOF'
## Summary

<what changed and why>

Closes #N

## Changes

- <bullet list of files or functions affected>

## Checklist

- [ ] `Invoke-Pester Tests/ -Output Detailed` passes with no failures
- [ ] PSScriptAnalyzer reports no Error or Warning violations on changed `.ps1` files
- [ ] No non-ASCII characters in changed `.ps1` files
- [ ] Commit messages follow Conventional Commits format
- [ ] `CHANGELOG.md` updated (if user-facing change)
- [ ] Documentation updated (README, CONTRIBUTING, or CLAUDE.md as applicable)
EOF
)"
```

Report the PR URL to the user and tell them to run `/merge-pr` once CI passes.
