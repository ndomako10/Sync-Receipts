Review all committed configuration templates for schema correctness and alignment with the code that reads them.

Read each of the following files in full:
- Config/Templates/Config.template.ini
- Config/Templates/Categories.template.json
- Config/Templates/Methods.template.json
- Scripts/Initialize-SyncReceipts.ps1 (template copy steps only)

Also run:
- `git grep -n "Get-Methods\|Get-Categories\|Get-ValidAccounts\|defaults\|InstallArgs\|template" Scripts/Sync-Receipts.ps1 Scripts/Initialize-SyncReceipts.ps1 Scripts/New-AccountsTemplate.ps1` to find where templates are referenced in code
- `ls Config/` to confirm which template files exist on disk
- `gh issue list --state open --label config` to check for known outstanding config gaps

Check each area and flag any issues:

1. **Config/Templates/Config.template.ini completeness** -- Does the template define every variable that the launchers and script use?
   - Required: `RECEIPTS_ROOT`, `WORKBOOKS_ROOT`, `DATE_FORMAT`
   - Optional but documented: `RECEIPTS_ROOT_LOCAL`
   - Flag any variable read by a launcher (`Launchers/*.bat`) or the script that is absent from the template, and any variable in the template that is no longer read anywhere.

2. **Config/Templates/Config.template.ini defaults** -- Are the documented default values consistent with the script's `param()` defaults?
   - `DATE_FORMAT=yyMMdd` must match the `$DateFormat` default in `Sync-Receipts.ps1`
   - Flag any mismatch.

3. **Categories.template.json schema** -- Validate the structure the code expects:
   - Root must be a JSON object (not an array)
   - Each key is a category display name (string); each value is an array of subcategory strings
   - No duplicate category keys
   - No duplicate subcategory strings within a category
   - Every category name and subcategory name must survive `ConvertTo-ExcelRangeName` without producing an empty string or a name starting with a digit after sanitisation (letters, digits, spaces, `&`, `/`, and `_` are the only characters seen in practice)
   - Flag any name that would produce a collision after sanitisation (e.g. two names that differ only by `&` vs space)

4. **Methods.template.json schema** -- Validate the structure the code expects:
   - Root must be a JSON array of strings
   - Each token must be a single word (letters, digits, underscore only -- no spaces or special characters), matching the validation in `Get-Methods`
   - `Cash` must NOT appear in the list (it is always valid implicitly and would be redundant)
   - The template tokens must match the built-in fallback defaults in `Get-Methods` (currently: Card, Check, Checking, Savings, Transfer, Wire). Flag any token present in one but not the other.

5. **Initialize-SyncReceipts.ps1 copy steps** -- For each template file, verify that the copy step in `Initialize-SyncReceipts.ps1` references the correct source filename and correct destination filename:
   - `Config/Templates/Categories.template.json` --' `Config/Categories.json`
   - `Config/Templates/Methods.template.json` --' `Config/Methods.json`
   - `Config/Templates/Accounts.template.xlsx` --' `Config/Accounts.xlsx`
   - Flag any mismatch between what the script copies and what actually exists on disk.

6. **Accounts.template.xlsx schema** -- The file cannot be read directly, but verify via `New-AccountsTemplate.ps1` that the column schema it writes matches what `Get-ValidAccounts` expects to read:
   - Columns (1-based): A=Last 4, B=Method, C=Holder, D=Institution, E=Account, F=Status
   - `Status` column must accept the values `Active` and `Inactive`
   - Flag any column name or position mismatch between the writer and the reader.

7. **Open issues** -- For each issue returned by `gh issue list --state open --label config`, state whether it is already addressed by the current templates (flag for closing if so) or remains outstanding.

## Output format

For each finding: `SEVERITY | FILE:LINE | one-sentence problem | one-sentence fix`
Severity: CRITICAL (broken/data-loss), MAJOR (correctness gap), MINOR (polish).
If no issues in an area: `OK: <area>`
No narrative. No file contents. Findings only.
