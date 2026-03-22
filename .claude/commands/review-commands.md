Review all slash commands in .claude/commands/ (except review-all.md) for accuracy, completeness, and token efficiency.

Run first:
- `ls .claude/commands/` to get the full file list
- `gh issue list --state open --json number,title,labels --limit 100`

Then read all command files found (except review-all.md) in parallel.

Check each command file:

1. **File path accuracy** --" every path referenced (files to read, scripts to run, dirs to list) exists in the repo; no hardcoded file list is missing a file that clearly belongs.

2. **Shell command accuracy** --" every `gh` label exists in `.github/labeling.yml`; every `git grep` pattern would produce results against the current codebase.

3. **Coverage completeness**
   - `review-tests.md`: covers every file in `Tests/`
   - `review-tooling.md`: covers every workflow in `.github/workflows/` and every hook in `Scripts/hooks/`
   - `review-config.md`: covers every template in `Config/Templates/`
   - `review-launchers.md`: covers every file in `Launchers/`

4. **Output contract** --" every command specifies a structured output format using CRITICAL/MAJOR/MINOR severity; no command instructs the agent to quote file contents instead of citing FILE:LINE.

5. **Token efficiency** --" flag:
   - Whole-file reads where a scoped read (section, offset/limit) would suffice
   - `content`-mode searches where `files_with_matches` suffices for existence checks
   - Verbose output contracts that allow prose, restated context, or full file contents
   - For each gap, propose specific replacement text.

6. **Open issues** --" for each open issue labelled `documentation`, `ci`, or `config`, check whether the relevant command's checklist covers it; flag any gap.

## Output format

For each finding:
  `SEVERITY | FILE:LINE | one-sentence problem | one-sentence fix`

Severity: CRITICAL (broken/data-loss), MAJOR (correctness gap), MINOR (polish).
If no issues in a command file: `OK: <filename>`
No narrative. No file contents. Findings only.
