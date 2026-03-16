# Known Issues

## #1 — Subcategory (column H) dropdown validation not implemented

**Status:** Blocked — pending implementation of zip binary patching approach

**Summary:** The Category dropdown (column G) works. The Subcategory dropdown (column H) — which should use `=INDIRECT(G2)` to show subcategories matching the selected category — does not work and is not currently set.

---

### What was tried and why it failed

#### Attempt 1 — COM `Validation.Add` with `INDIRECT`

```powershell
$range.Validation.Add(3, 1, 1, "=INDIRECT(G2)")
```

Excel throws `0x800A03EC` when any formula-based list validation using `INDIRECT` is set via COM. This is a known Excel restriction with no workaround via COM alone.

#### Attempt 2 — XML injection via `ZipArchive.Update`

Wrote `dataValidations` XML directly into the sheet's zip entry after Excel saves and closes. Failed because `.NET ZipArchive` (Update mode) writes patched entries as **STORED** (uncompressed). Excel requires **DEFLATED**.

#### Attempt 3 — XML injection via `ZipArchive.Create` (read-all-into-memory)

Switched to reading all entries into memory and writing a new archive with `CompressionLevel.Optimal`. Produced DEFLATED entries but Excel still rejected the file.

**Root cause identified:** Excel requires specific values in the zip headers that .NET `ZipArchive` does not write:
- Local file headers (`PK\x03\x04`): `flag_bits` must be `0x0006`; .NET writes `0x0000`
- Central directory headers (`PK\x01\x02`): `version_made_by` must be `45` (0x2D); .NET writes `20`

---

### What needs to be implemented

A function `Set-SubcategoryValidationXml` that runs **after** `$workbook.Save()` and `$workbook.Close()`. Steps:

1. Read all zip entries into memory (`ZipFile.OpenRead`)
2. Resolve sheet name → zip path via `workbook.xml` + `workbook.xml.rels` using XML DOM (not regex — all `<sheet>` elements are on one line and regex will match incorrectly across them)
3. Patch the target sheet XML:
   - Remove any existing `<dataValidations>` element
   - Inject a new one with `count="2"` containing both Category and Subcategory validation
   - Generate a random GUID for each `xr:uid` attribute (the `xr:` namespace is already declared on `<worksheet>`)
4. Write a fresh zip using `ZipFile.Open` in Create mode
5. **After closing the ZipArchive**, scan the output binary and patch every local file header and central directory header (see byte offsets below)
6. Replace the original file with the patched temp file (verify temp file size > 1KB first)

> **Note:** Because both validations are injected via XML, the existing `Validation.Add` COM call for Category in `Sync-Month` should be removed to avoid conflicts.

---

### Target XML

```xml
<dataValidations count="2">
  <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"
    sqref="G2:G{lastRow}" xr:uid="{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}">
    <formula1>Category</formula1>
  </dataValidation>
  <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"
    sqref="H2:H{lastRow}" xr:uid="{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}">
    <formula1>INDIRECT(G2)</formula1>
  </dataValidation>
</dataValidations>
```

---

### Binary patch details

#### Local file headers — signature `PK\x03\x04` (50 4B 03 04)

| Offset | Size | Field | Action |
|--------|------|-------|--------|
| 0 | 4 | Signature | Do not change |
| 4 | 2 | Version needed | Leave as-is |
| 6 | 2 | General purpose bit flag | Set to `06 00` (0x0006) |
| 8 | 2 | Compression method | Leave as-is |

#### Central directory headers — signature `PK\x01\x02` (50 4B 01 02)

| Offset | Size | Field | Action |
|--------|------|-------|--------|
| 0 | 4 | Signature | Do not change |
| 4 | 2 | Version made by | Set to `2D 00` (45, Windows) |
| 6 | 2 | Version needed | Leave as-is |
| 8 | 2 | General purpose bit flag | Set to `06 00` (0x0006) |
