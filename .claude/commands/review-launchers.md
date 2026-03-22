Review all launcher batch files for correctness and consistency.

Read each of the following files in full:
- Launchers/Run-SyncReceipts.bat
- Launchers/Run-SyncMonthReceipts.bat
- Launchers/Run-SyncYearReceipts.bat
- Launchers/Run-SyncAllReceipts.bat
- Config/Templates/Config.template.ini

Also run:
- `git grep -n "^param" Scripts/Sync-Receipts.ps1` to locate the param block, then read those lines to get the current parameter names and defaults
- `gh issue list --state open --label config` to check for known outstanding launcher or config gaps

Check each area and flag any issues:

1. **Parameter alignment** -- For each launcher, verify that every parameter passed to `Sync-Receipts.ps1` matches an actual parameter in its `param()` block (correct spelling, correct switch vs string type). Flag any parameter name that does not match. Also check that no launcher omits a required parameter.
   - All launchers must pass: `-ReceiptsRoot`, `-WorkbooksRoot`, `-DateFormat`
   - `Run-SyncMonthReceipts.bat` must pass `-YearMonth`
   - `Run-SyncYearReceipts.bat` must pass `-Year`
   - `Run-SyncAllReceipts.bat` must pass `-All`

2. **Config.ini reading pattern** -- All launchers must use the same `findstr /v "^#"` pattern to read `Config\Config.ini`. Flag any launcher that uses a different approach or a different relative path to locate Config.ini.

3. **Variable names** -- Variables read from Config.ini (`RECEIPTS_ROOT`, `WORKBOOKS_ROOT`, `DATE_FORMAT`) must match the variable names defined in `Config/Templates/Config.template.ini`. Flag any mismatch.

4. **WORKBOOKS_ROOT fallback** -- Every launcher must include `if "%WORKBOOKS_ROOT%"=="" set "WORKBOOKS_ROOT=%RECEIPTS_ROOT%"` immediately after reading Config.ini. Flag any launcher that omits it or places it elsewhere.

5. **DATE_FORMAT default** -- The hardcoded `set "DATE_FORMAT=yyMMdd"` must appear before the Config.ini read in every launcher (so Config.ini can override it), and the default value `yyMMdd` must match the default in `Sync-Receipts.ps1`'s `param()` block.

6. **UNC path handling** -- Every launcher must use `pushd "%~dp0"` at the top and `popd` at the bottom. Flag any launcher that omits either.

7. **Script path** -- Every launcher must reference `"%~dp0..\Scripts\Sync-Receipts.ps1"`. Flag any launcher using an absolute path, a different relative path, or a different script name.

8. **Consistency** -- All launchers should share the same structure (comment header, pushd, DATE_FORMAT default, Config.ini read, WORKBOOKS_ROOT fallback, PowerShell invocation, pause, popd). Flag any structural deviation that is not explained by the launcher's specific purpose.

9. **Completeness** -- Are there any sync modes supported by `Sync-Receipts.ps1` (via `-YearMonth`, `-Year`, `-All`, or default current-month) that lack a corresponding launcher? Flag gaps.

## Output format

For each finding: `SEVERITY | FILE:LINE | one-sentence problem | one-sentence fix`
Severity: CRITICAL (broken/data-loss), MAJOR (correctness gap), MINOR (polish).
If no issues in an area: `OK: <area>`
No narrative. No file contents. Findings only.
