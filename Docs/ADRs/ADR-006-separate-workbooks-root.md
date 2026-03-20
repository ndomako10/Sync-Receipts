# ADR-006: Add WorkbooksRoot Parameter to Separate Workbook Output from Receipts

## Status
Accepted

## Context
Receipts live on a network share (`ReceiptsRoot`). Writing the per-year workbooks
to the same share means every sync involves large .xlsx reads and writes over the
network, which is slow and risks corruption if the connection drops mid-write.
(Issue #31.)

## Decision
Add a `-WorkbooksRoot` parameter to `Sync-Receipts.ps1` (and `WORKBOOKS_ROOT` to
`Config.template.env`). When set, workbooks are written to that directory instead
of `ReceiptsRoot`. Defaults to `ReceiptsRoot` when not provided, preserving existing
behaviour.

## Alternatives considered
- **Always write workbooks to a fixed local path** -- simpler, no parameter needed.
  Rejected because the local path varies by machine and hardcoding it would break
  portability across machines sharing the same `Config.template.env`.
- **Write workbooks to the repo directory** -- keeps output near the script. Rejected
  because the repo directory may be on a different drive or a dev machine, and mixing
  generated output with source files is undesirable.
- **Cache locally, sync to share on close** -- more robust but requires a separate
  sync step and conflict resolution if the share copy diverges. Over-engineered for
  a single-user tool.

## Consequences
- Users with receipts on a network share can set `WORKBOOKS_ROOT` to a local drive
  for fast writes; workbooks remain accessible on the local machine.
- Two directory variables must be kept consistent in `Config.env`; setup documentation
  must explain both.
- `WorkbookPath` (for test overrides) still takes precedence over both roots when set.
