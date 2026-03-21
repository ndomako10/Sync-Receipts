# ADR-008: Use .env Format for Configuration File Instead of Batch

## Status
Superseded by ADR-011

Superseded by [ADR-011](ADR-011-ini-config-format.md) which documents the rename to .ini format.

## Context
The script requires one machine-specific setting at minimum: `RECEIPTS_ROOT`, the path
to the user's receipt folder. This value had to be stored somewhere the batch launchers
could read it before invoking PowerShell.

The original implementation used a Windows batch file (`Config.bat`) with `set` syntax:

    set "RECEIPTS_ROOT=C:\Users\...\Receipts"

This was functional but required the user to edit a file containing batch syntax. For
non-technical users, the `set "KEY=value"` form is unfamiliar and error-prone -- a
misplaced quote or trailing space silently corrupts the value. The batch file also
ran in-process in the launcher (`call Config.bat`), meaning a syntax error in Config.bat
would abort the entire launcher without a clear error message.

## Decision
Replace `Config.bat` with `Config.env`, a plain `KEY=value` text file with `#`-prefixed
comment lines. Launchers parse it with a `for /f` + `findstr` one-liner:

    for /f "usebackq tokens=1,* delims==" %%A in (`findstr /v "^#" "%~dp0..\Config\Config.env"`) do
        if not "%%A"=="" set "%%A=%%B"

This keeps all configuration in a single, clearly structured file that any text editor
can open without knowledge of batch syntax.

## Alternatives considered
- **Keep Config.bat with `set` syntax** -- zero parser complexity; the file simply runs
  in-process. Rejected because the `set "KEY=value"` form is unfamiliar to non-technical
  users and quoting rules are non-obvious. A bare `=` in a value, a trailing space, or a
  missing quote produces silent corruption.
- **Use a PowerShell `.ps1` config file** -- idiomatic for a PowerShell project, but the
  launchers are batch files that must pass arguments to PowerShell; reading a `.ps1` from
  batch requires an extra `powershell -Command` invocation before the main script call,
  adding complexity and a second PowerShell startup.
- **Pass parameters interactively** -- avoids a config file entirely. Rejected because
  the primary use case is double-clicking a shortcut with no user interaction.

## Consequences
- Each launcher must include the `for /f` parser instead of a single `call` line.
  The parser is slightly more complex but self-contained and well-understood.
- Blank values (`KEY=`) are handled naturally: `for /f` skips empty lines; a blank
  value sets the variable to an empty string, which the script interprets as "not set"
  and applies its default.
- Comment lines (starting with `#`) are stripped by `findstr /v "^#"` before parsing.
- Existing users upgrading from `Config.bat` must copy their values into the new
  `Config.env`. `Initialize-SyncReceipts.ps1` was updated to create `Config.env`
  from `Config.template.env`. The old `Config.bat` path is retained in `.gitignore`
  to avoid surfacing it as an untracked file for users mid-upgrade.
