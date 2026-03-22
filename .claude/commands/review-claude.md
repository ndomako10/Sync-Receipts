Review the CLAUDE.md files for accuracy, completeness, and token efficiency.

Read in parallel:
- `CLAUDE.md`
- `~/.claude/CLAUDE.md`

Run in parallel:
- `git grep -n "^function " Scripts/Sync-Receipts.ps1`
- `ls Docs/ADRs/`
- `ls .claude/commands/`
- `gh issue list --state open --json number,title --limit 100`

Check each area:

1. **Architecture tree** --" every path in the tree exists on disk; every committed file/folder appears in the tree; gitignored annotations are correct.

2. **Key Functions table** --" every `function` from `git grep` is in the table with an accurate description; no table entry is missing from the script.

3. **Commit scope table** --" every scope row references a file/area that still exists; no significant file/area lacks a scope.

4. **Coding rules** --" every referenced function, pattern, or file still exists; no rule contradicts the current codebase; no recurring codebase pattern is undocumented.

5. **Global/project consistency** --" project CLAUDE.md overrides rather than duplicates global rules; no conflict without an explicit override note.

6. **ADR index** --" every file in `Docs/ADRs/` appears in `Docs/ADRs/README.md` and vice versa; significant undocumented decisions flagged.

7. **Slash command index** --" every file in `.claude/commands/` is referenced in CLAUDE.md or CONTRIBUTING.md.

8. **Token efficiency** --" flag any instruction that directs agents to read more than needed (whole file when a section suffices, `content` mode when `files_with_matches` suffices, re-reading files already in context, verbose agent output where compact suffices).

## Output format

For each finding:
  `SEVERITY | FILE:LINE | one-sentence problem | one-sentence fix`

Severity: CRITICAL (broken/data-loss), MAJOR (correctness gap), MINOR (polish).
If no issues in an area: `OK: <area>`
No narrative. No file contents. Findings only.
