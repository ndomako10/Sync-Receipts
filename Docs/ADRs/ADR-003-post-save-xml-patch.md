# ADR-003: Inject Dropdown Validation via Post-Save XML Patch

## Status
Accepted

## Context
The workbook requires two dependent dropdown columns: Category (free list) and
Subcategory (filtered by the selected Category). Excel's native dependent validation
requires a formula like `=INDIRECT(Category)`, which Excel's COM object model cannot
write to a hidden sheet -- attempting to do so causes a COM deadlock or silent failure.

## Decision
Write the workbook to disk via COM, then reopen the .xlsx as a ZIP archive using
`System.IO.Compression.ZipFile`, locate `xl/worksheets/sheet*.xml`, and inject the
`<dataValidation>` elements directly into the XML before rewriting the ZIP entry.
A binary header fixup is applied after ZIP rewrite because `ZipFile` overwrites the
first four bytes of the ZIP local file header.

This is implemented in `Set-SubcategoryValidationXml`.

## Alternatives considered
- **COM-based validation on visible sheets** -- works for simple list validation on
  visible sheets, but fails silently or deadlocks when targeting hidden sheets (the
  Category sheet is hidden to keep the workbook clean).
- **Unhide the Category sheet, apply validation via COM, rehide** -- attempted; still
  produced intermittent COM hangs during named range manipulation on the hidden sheet
  (issue #22). Unreliable.
- **Store validation ranges on a visible sheet** -- would expose internal category
  data to users, cluttering the workbook. Rejected for UX reasons.
- **Use a macro / VBA** -- requires enabling macros, which is a security barrier for
  most users and cannot be distributed as a plain .xlsx.

## Consequences
- The patch runs after every save, adding a fixed overhead per sync.
- The binary header fixup is fragile: it assumes the ZIP local file header magic
  bytes are at offset 0. If `ZipFile` behaviour changes across .NET versions this
  could corrupt the file.
- The approach is not portable to workbooks opened by other tools that rewrite the
  ZIP structure (e.g. LibreOffice resave).
- Debugging validation failures requires inspecting the raw XML inside the .xlsx ZIP.
