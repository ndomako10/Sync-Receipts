Review all PowerShell comment-based help for accuracy and completeness.

Run first to locate comment-based help blocks without reading full files:
- `git grep -n "^<#\|^\.SYNOPSIS\|^\.DESCRIPTION\|^\.PARAMETER\|^\.EXAMPLE\|^#>" Scripts/Sync-Receipts.ps1 Scripts/Initialize-SyncReceipts.ps1 Scripts/Install-GitHooks.ps1 Scripts/New-AccountsTemplate.ps1`
- `git grep -n "^function " Scripts/Sync-Receipts.ps1 Scripts/Initialize-SyncReceipts.ps1 Scripts/Install-GitHooks.ps1 Scripts/New-AccountsTemplate.ps1`

Then read only the help sections identified above (using offset/limit) plus the `param()` block of each script. Do NOT read these files in full; use targeted reads.

Check each function and script against the actual code and flag any issues in these areas:

1. **Script-level help** -- Does each script have a `.SYNOPSIS`, `.DESCRIPTION`, and `.EXAMPLE`? Does `.DESCRIPTION` accurately list all setup steps or behaviours? Are `.PARAMETER` blocks present for every parameter in the `param()` block?

2. **Function-level help** -- For each `function` definition:
   - Does it have at least `.SYNOPSIS` and `.PARAMETER` blocks?
   - Do `.PARAMETER` descriptions match the actual parameter names and types in the `param()` block?
   - Are any parameters documented that no longer exist, or any parameters missing documentation?
   - Does `.SYNOPSIS` accurately describe what the function does?

3. **Examples** -- Are `.EXAMPLE` blocks present and accurate? Do they use realistic values consistent with the project (e.g. correct paths, valid method tokens, realistic account numbers)?

4. **Stale references** -- Any help text referencing removed parameters, old file names, or outdated behaviour?

5. **Cross-reference with open issues** -- Run `gh issue list --state open --label documentation` and check whether any open documentation issues are already addressed by the current help text (flag for closing) or remain outstanding.

## Output format

For each finding: `SEVERITY | FILE:LINE | one-sentence problem | one-sentence fix`
Severity: CRITICAL (broken/data-loss), MAJOR (correctness gap), MINOR (polish).
If no issues in an area: `OK: <area>`
No narrative. No file contents. Findings only.
