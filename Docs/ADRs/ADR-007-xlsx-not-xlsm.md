# ADR-007: Keep Workbooks as .xlsx (Non-Macro-Enabled)

## Status
Accepted

## Context
Several desirable features require event-driven logic that runs inside Excel as the
user edits cells -- clearing a stale subcategory when the category changes (issue #33),
protecting script-owned columns from accidental edits, real-time data validation, and
a one-click re-sync button. All of these require VBA, which Excel only supports in
macro-enabled workbooks (.xlsm).

The current workbooks are plain .xlsx. Switching to .xlsm would require:
- Renaming every per-year workbook (2025.xlsx -> 2025.xlsm, etc.) -- a breaking change
  for existing users
- Embedding a binary vbaProject.bin in the post-save XML patch (ADR-003), which is
  significantly more complex than patching plain XML
- Users receiving an "Enable Content" security prompt every time they open a workbook

## Decision
Keep workbooks as plain .xlsx. Features that would otherwise require VBA event
handlers are implemented as sync-time logic instead -- validated and corrected on
the next sync run rather than in real time.

Issue #33 (clear subcategory when category changes) is implemented by validating
the preserved Subcategory value against the category list during Write-MonthSheet:
if the Subcategory is not in the list for the current Category, it is written back
as blank.

## Alternatives considered
- **Switch to .xlsm** -- unlocks Worksheet_Change events, cell protection, real-time
  validation, and in-workbook automation buttons. Rejected because the binary VBA
  project is difficult to version and review, the format change is breaking for
  existing users, and the macro security prompt adds friction. Tracked as a future
  option if real-time interaction becomes a priority.

## Consequences
- Data quality features (invalid subcategory, stale values) are caught on re-sync,
  not immediately. A stale subcategory persists in the workbook until the next sync.
- The post-save XML patch (ADR-003) remains plain XML manipulation -- no binary
  embedding required.
- If a future issue requires real-time Excel interaction, switching to .xlsm will
  be a breaking change and should be tracked as a separate issue with its own ADR.
