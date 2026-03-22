# /merge-pr

Merge the current branch's PR once CI passes, then clean up the local branch.

Run this command after `/prep-pr` has opened the PR and CI has had time to run.

## Step 1 -- Get PR status

```bash
gh pr view --json number,title,url,statusCheckRollup,mergeable
```

## Step 2 -- Check CI

Inspect `statusCheckRollup`:
- If any check is `PENDING` or `IN_PROGRESS`: tell the user CI is still running and to
  try again in a moment. Stop here.
- If any check has `FAILURE` or `ERROR`: report which checks failed and stop. Do not merge.
- If all checks are `SUCCESS`: proceed.

Also check `mergeable`:
- If `CONFLICTING`: tell the user to resolve conflicts first. Stop here.

## Step 3 -- Confirm merge

Show the user:
- PR number and title
- All CI checks passed
- "Ready to merge?"

Wait for confirmation before proceeding.

## Step 4 -- Merge

```bash
gh pr merge --merge --delete-branch
```

Use `--merge` (not squash or rebase) to preserve the commit history.
`--delete-branch` removes the remote branch (local cleanup follows below).

## Step 5 -- Clean up local branch

```bash
git checkout master
git pull origin master
git branch -d <branch-name>
```

Report that the PR is merged and that `auto-release.yml` will cut a release
automatically if the CHANGELOG had an `[Unreleased]` entry.
