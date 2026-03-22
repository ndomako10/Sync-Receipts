Review all project documentation for accuracy and completeness.

Read each of the following files in full:
- README.md
- CLAUDE.md
- CONTRIBUTING.md
- SECURITY.md
- Config/Templates/Config.template.ini

Also run:
- `git grep -n "^function " Scripts/Sync-Receipts.ps1` to get all function names and line numbers for the cross-reference check
- `gh issue list --state open --label documentation` to check for known outstanding doc gaps

Check each doc against the actual code and flag any issues in these areas:

1. **Architecture** --" Does the script/data decoupling description match reality?
   - PS1 scripts live in Scripts/, batch launchers in Launchers/, tests in Tests/, ADRs in Docs/ADRs/
   - Receipt folders live at RECEIPTS_ROOT; per-year workbooks are written to WORKBOOKS_ROOT (defaults to RECEIPTS_ROOT when not set)
   - `-ReceiptsRoot` must always be provided explicitly via Config/Config.ini
   - Categories.json lives in Config/ and is read via Join-Path (Split-Path $PSScriptRoot -Parent) "Config"

2. **Function table (CLAUDE.md)** --" Use the `git grep` output above to get all function names,
   then cross-reference against the Key Functions table in CLAUDE.md. Flag any function that is
   defined in the script but missing from the table, or listed in the table but no longer exists
   in the script. Verify that descriptions are accurate for functions that are present in both.

3. **Parameters** --" Do the documented parameters match the actual `param()` block in Scripts/Sync-Receipts.ps1?

4. **File/folder structure** --" Are all committed files and folders (Config/, Scripts/, Launchers/, Tests/, Docs/ADRs/) reflected accurately in README.md and CLAUDE.md?

5. **Config/Templates/Config.template.ini** --" Does it list all variables that Config.ini should define?

6. **Scope table (CLAUDE.md and CONTRIBUTING.md)** --" Do the commit scope definitions match the actual files and areas in the repo?

7. **Stale references** --" Any mentions of removed features, old file names (e.g. Config.bat, Config.template.bat, lowercase paths), or outdated patterns?

8. **ADRs** --" Run `ls Docs/ADRs/` to list all ADR files on disk. Check:
   - Every ADR file is indexed in Docs/ADRs/README.md (no orphaned files, no missing entries)
   - No ADR references a file, function, or parameter that no longer exists in the codebase
   - Are there significant decisions made recently that warranted an ADR but have none? (Check recent commits and issues for choices between two or more reasonable alternatives with non-obvious tradeoffs)

## Output format

For each finding: `SEVERITY | FILE:LINE | one-sentence problem | one-sentence fix`
Severity: CRITICAL (broken/data-loss), MAJOR (correctness gap), MINOR (polish).
If no issues in an area: `OK: <area>`
No narrative. No file contents. Findings only.
