# Sync-Receipts.ps1 — Session Handoff

## What the script does

Syncs receipt files from a folder structure into a pre-existing Excel workbook (`Receipts.xlsx`).
Each receipt filename encodes its metadata: `YYMMDD Vendor -$Amount Method [Account]`
The script parses each filename, writes a row to a sheet named in YYMM format (e.g. `2603`),
creates a formatted table, validates accounts, and sets category dropdowns.

## Folder/file structure

```
\\YOUR-PC\Shared\Documents\Receipts\
    Receipts.xlsx              ← target workbook (must be named exactly this)
    Sync-Receipts.ps1
    Run-SyncReceipts.bat
    Run-SyncAllReceipts.bat
    2026\
        2603 - March\
            260301 Vendor -$10.00 Card 1234.pdf
            ...
```

## Workbook structure

- **Account sheet** — column A, row 2 down: 4-digit account numbers used for validation
- **Category sheet** — row 1 is category headers, rows 2+ are subcategories per column
- **Named ranges** — one per category (e.g. `Housing`, `Food`) pointing to subcategory lists; `Category` range points to header row. Set by script via COM each run.
- **Month sheets** (e.g. `2603`) — created/overwritten by the script. 9 columns: File Name, Date, Vendor, Amount, Method, Account, Category, Subcategory, Flag

## Script parameters

| Parameter | Description |
|-----------|-------------|
| `-ReceiptsRoot` | Path to root Receipts folder. Defaults to script's own folder. |
| `-YearMonth` | YYMM to sync (e.g. `2603`). Defaults to current month. |
| `-All` | Sync every month folder found under every year folder. |
| `-KillExcel` | Kill any running EXCEL.EXE before starting. Use when Excel crashed and locked the file. |

## Current state of the script

The script is **fully working except for subcategory (column H) dropdown validation.**

- Category dropdown (column G) works via COM (`=Category` named range)
- Subcategory dropdown (column H) is **not currently set** — this was stripped out pending a solution

### What was tried for subcategory validation and why it failed

**COM `Validation.Add` with `INDIRECT`** — Excel throws `0x800A03EC` when setting any formula-based list validation that uses `INDIRECT` via COM. This is a known Excel restriction.

**XML injection (multiple attempts)** — The approach was to write the `dataValidations` XML directly into the sheet's zip entry after Excel saves and closes. This failed due to a series of zip format issues:

1. `.NET ZipArchive.Update` writes patched entries as **STORED** (uncompressed). Excel requires **DEFLATED**. This caused corruption.
2. Switched to read-all-into-memory + `ZipArchive.Create` with `CompressionLevel.Optimal`. This produced DEFLATED entries but Excel still rejected the file.
3. Root cause identified by comparing working vs broken files: Excel requires zip local file header **`flag_bits = 0x0006`** and **`create_version = 45`** on every entry. .NET `ZipArchive` writes `flag_bits = 0x0000` and `create_version = 20`. Excel rejects entries without these flags.

### What needs to be implemented

Re-add `Set-SubcategoryValidationXml` and a Phase 2 block. The function must:

1. Read all zip entries into memory (using `ZipFile.OpenRead`)
2. Resolve the sheet name → zip path via `workbook.xml` and `workbook.xml.rels` using XML DOM (not regex — all `<sheet>` elements are on one line and regex will match across them)
3. Patch the target sheet XML: remove any existing `<dataValidations>` element, inject a new one with **both** Category and Subcategory validation together (count="2"), each with a generated GUID for `xr:uid`
4. Write a fresh zip using `ZipFile.Open` in Create mode
5. **After closing the ZipArchive**, patch the raw binary of the output file to set `flag_bits = 0x0006` and `create_version = 45` on every local file header entry. This must be done by scanning the binary for the local file header signature `PK\x03\x04` and patching the correct byte offsets.
6. Replace the original file with the patched temp file, verifying temp file size > 1KB first

### The working dataValidations XML (from a known-good file)

```xml
<dataValidations count="2">
  <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"
    sqref="G2:G36" xr:uid="{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}">
    <formula1>Category</formula1>
  </dataValidation>
  <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"
    sqref="H2:H36" xr:uid="{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}">
    <formula1>INDIRECT(G2)</formula1>
  </dataValidation>
</dataValidations>
```

The `xr:uid` GUIDs can be randomly generated per run — they just need to be present and in valid GUID format. The `xr:` namespace is already declared on the `<worksheet>` root element.

Note: because both validations are injected via XML, the COM `Validation.Add` call for Category in `Sync-Month` should be **removed** to avoid conflicts.

### Local file header binary patch details

The zip local file header format starting at each `PK\x03\x04` signature:

| Offset | Size | Field | Target value |
|--------|------|-------|--------------|
| 0 | 4 | Signature | `50 4B 03 04` (do not change) |
| 4 | 2 | Version needed to extract | leave as-is |
| 6 | 2 | **General purpose bit flag** | set to `06 00` (0x0006) |
| 8 | 2 | Compression method | leave as-is |
| ... | | | |

`create_version` (45) lives in the **central directory** headers, not local headers. Central directory entries start with `PK\x01\x02`:

| Offset | Size | Field | Target value |
|--------|------|-------|--------------|
| 0 | 4 | Signature | `50 4B 01 02` |
| 4 | 2 | **Version made by** | set to `2D 00` (45, Windows) |
| 6 | 2 | Version needed | leave as-is |
| 8 | 2 | **General purpose bit flag** | set to `06 00` |

Both local and central directory flag bytes need to be patched.

## Coding rules established this session

- **Always add error handling and debug output when making changes** — every new block needs try/catch and Write-Host logging
- **Never use `$variable:` in double-quoted strings** — PowerShell interprets the colon as a drive separator. Use `${variable}:` instead
- **No smart quotes or em-dashes** — the file must be pure ASCII. Any non-ASCII characters break PowerShell parsing on the network share
- **Always verify non-ASCII bytes = 0 after any edit** before copying to outputs
- **Propose changes before making them** — do not edit code without confirmation

## Key files

- `Sync-Receipts.ps1` — main script (current working version attached to this session)
- `Run-SyncReceipts.bat` — runs script for current month against the path in `config.bat`
- `Run-SyncAllReceipts.bat` — runs script with `-All` flag
- `Deploy-SyncReceipts.bat` — machine-specific deployment helper (not committed)
- `config.bat` — local machine settings (not committed; copy from `config.template.bat`)
