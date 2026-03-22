Run all review checks, fix issues inline, create GitHub issues for significant findings, and report back.

---

## Phase 1 --" Fix CLAUDE.md (foreground agent)

Spawn a foreground Agent with this prompt:

> Read `.claude/commands/review-claude.md` and follow its instructions. Apply every CRITICAL and MAJOR fix directly to the affected files using the Edit tool. Return only:
> - `FIXED | FILE | change` for each fix applied
> - `MINOR | FILE:LINE | issue` for unfixed minor findings
> - `OK: <area>` for clean areas

Wait for completion before proceeding. Record the output.

---

## Phase 2 --" Fix slash commands (foreground agent)

Spawn a foreground Agent with this prompt:

> Read `.claude/commands/review-commands.md` and follow its instructions. Apply every CRITICAL and MAJOR fix directly to the affected files using the Edit tool. Return only:
> - `FIXED | FILE | change` for each fix applied
> - `MINOR | FILE:LINE | issue` for unfixed minor findings
> - `OK: <area>` for clean areas

Wait for completion before proceeding. Record the output.

---

## Phase 3 --" Run all review commands (3 background agents)

In a **single message**, launch all 3 agents with `run_in_background: true`.

Each agent fetches its own data --" nothing is pre-fetched here. Each agent creates GitHub issues for CRITICAL/MAJOR findings not already in the tracker, then returns only URLs and clean-area signals.

**Issue body template** (all agents use this):
```
## Finding
<one paragraph: what is wrong and where>
## Impact
<one sentence>
## Fix
<one sentence or short list>
```

**Output contract** (append to all 3 agent prompts):
> Return only: `CREATED #N <title>` for each issue created; `EXISTS #N` for already-tracked findings; `OK: <area>` for clean areas. No narrative, no findings text, no file contents.

---

### Agent 1 prompt --" Docs --- Help --- Changelog

```
Run `gh issue list --state open --json number,title,labels --limit 100` first.

Read and execute the review instructions from these files in order:
- .claude/commands/review-docs.md
- .claude/commands/review-help.md
- .claude/commands/review-changelog.md

For each CRITICAL or MAJOR finding not already tracked, create a GitHub issue using the issue body template. Use label: documentation.

[INSERT OUTPUT CONTRACT]
```

### Agent 2 prompt --" Config --- Launchers --- Security

```
Run `gh issue list --state open --json number,title,labels --limit 100` first.

Read and execute the review instructions from these files in order:
- .claude/commands/review-config.md
- .claude/commands/review-launchers.md
- .claude/commands/review-security.md

For each CRITICAL or MAJOR finding not already tracked, create a GitHub issue using the issue body template. Use label matching the area: config, security, setup.

[INSERT OUTPUT CONTRACT]
```

### Agent 3 prompt --" Tests --- Tooling

```
Run `gh issue list --state open --json number,title,labels --limit 100` first.

Read and execute the review instructions from these files in order:
- .claude/commands/review-tests.md
- .claude/commands/review-tooling.md

For each CRITICAL or MAJOR finding not already tracked, create a GitHub issue using the issue body template. Use label matching the area: test, ci.

[INSERT OUTPUT CONTRACT]
```

---

## Phase 4 --" Final summary

Wait for all 3 background agents to complete, then print:

```
## Fixes applied (Phases 1--"2)
- FILE | change

## Issues created
- #N title (area)

## Already tracked
- #N matches finding

## Clean areas
- OK: review-changelog

## Minor findings (not actioned)
- FILE:LINE | issue
```
