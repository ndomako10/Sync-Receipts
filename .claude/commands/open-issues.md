Fetch and analyse all open GitHub issues for this repo, then produce a structured report.

## Step 1 -- Fetch issues

Run this command to get all open issues with their labels and body text:

```bash
gh issue list --repo "$(git remote get-url origin)" --state open --json number,title,labels,body,createdAt,updatedAt
```

## Step 2 -- Analyse each issue

For each issue, infer the following fields from its title, body, and labels:

| Field | Values | How to infer |
|---|---|---|
| **Area** | Core script, Configuration, Tests, CI/Tooling, Docs/ADRs, Launchers | What part of the codebase does it touch? |
| **Type** | `feat`, `fix`, `docs`, `test`, `refactor`, `chore`, `ci` | What kind of change is needed? (Conventional Commits types) |
| **Difficulty** | Easy, Medium, Hard | Easy = well-scoped, no risky side-effects; Hard = broad, complex, or uncertain |
| **Risk** | Low, Medium, High | Could this break existing behaviour or require careful migration? |

## Step 3 -- Group and display

Group the issues by **Area**. Within each area, render a markdown table:

```
### Core script

| # | Title | Type | Difficulty | Risk |
|---|-------|------|------------|------|
| #12 | Fix amount parsing for refunds | fix | Easy | Low |
| #7  | Support multi-page receipts | feat | Hard | Medium |
```

Repeat for each area that has at least one issue.

## Step 4 -- Summary

After all tables, print a brief summary:

```
### Summary
- Total open: N  (X fixes - Y features - Z docs - W tests - ...)
- [One sentence per area that has issues]
```

## Step 5 -- Branch plan

Group the issues into **waves** -- ordered sets of branches that can be worked on in sequence. Within a wave, branches are independent of each other and can be done in parallel or any order. Across waves, later waves depend on earlier ones being merged first.

Use this logic to assign issues to waves:

1. **Wave 1 -- Bugs first.** All `fix`-type issues go here. They are small, unblock everything else, and should be resolved before adding new behaviour on top of broken code. Each fix gets its own branch.
2. **Wave 2 -- Docs and chores.** Pure `docs` or `chore` issues. Each gets its own branch and PR; no formal review required.
3. **Wave 3 -- Test cleanup.** Easy, self-contained `test` issues with no dependencies. Batch related test files into one branch where it avoids nearly-identical PRs.
4. **Wave 4 -- Prerequisite features.** `feat` issues that other issues depend on (e.g. a parameter that an integration test will use). Sequence them: note which must land before the next branch starts.
5. **Wave 5+ -- Independent features.** Remaining self-contained features, one branch each, ordered easy-before-hard.
6. **Final wave -- Large or blocked features.** Hard, high-risk, or design-uncertain issues. Note what each one depends on.

Render each wave as a table:

```
### Wave 1 -- Bug fixes

| Branch | Issues | Notes |
|---|---|---|
| `fix/65-install-module-splatting` | #65 | Broken setup; fix first so fresh installs work |
| `fix/66-numberformat-empty-month` | #66 | One guard clause in Write-MonthSheet |

### Wave 2 -- Docs and chores

| Branch | Issues | Notes |
|---|---|---|
| `docs/78-consolidate-initialize-help` | #78 | No formal review required |

### Wave 3 -- Test cleanup

| Branch | Issues | Notes |
|---|---|---|
| `test/74-75-77-test-suite-cleanup` | #74, #75, #77 | Same area; batch to avoid near-identical PRs |
```

Continue for all waves. For dependency chains within a wave, add a **Depends on** column and note which branch must merge first.
