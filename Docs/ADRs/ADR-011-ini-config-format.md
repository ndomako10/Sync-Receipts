# ADR-011: Rename Configuration File Extension from .env to .ini

## Status
Accepted

## Context
The project configuration file (`Config/Config.env`) uses a plain `KEY=value` format with
`#`-prefixed comments. The `.env` extension was adopted in ADR-008 as a step up from the
original `Config.bat` batch syntax.

Two practical problems emerged with `.env`:

1. **Windows file association**: Windows has no built-in association for `.env`. Double-
   clicking the file produces an "How do you want to open this file?" prompt, which is
   confusing for non-technical users who need to edit their configuration.
2. **Editor syntax highlighting**: Code editors (VS Code, Notepad++) have built-in grammar
   support for `.ini` files but not for `.env` files. The `.ini` extension gets colour
   coding automatically; `.env` does not.

The `.env` extension also implies Node.js / Docker tooling semantics (`dotenv` library),
which is not the context here. The file has no shell expansion, no quoting rules, and no
export semantics -- it is a simple `KEY=value` file closer to INI convention.

## Decision
Rename `Config/Config.env` to `Config/Config.ini` and `Config/Templates/Config.template.env`
to `Config/Templates/Config.template.ini`. The file format (KEY=value, # comments, parsed
by a batch `for /f` + `findstr` one-liner) is unchanged.

## Alternatives considered
- **Keep .env** -- avoids a breaking change. Rejected because the Windows file-association
  prompt is a real friction point for the target user (non-technical, double-clicks files
  to edit them).
- **Use .txt** -- universally recognised, no syntax highlighting. Rejected because `.ini`
  carries the correct semantic meaning and is supported by editors.
- **Use .ps1 config** -- idiomatic for PowerShell, but requires an extra PowerShell startup
  from the batch launchers. Rejected (same reasoning as ADR-008).

## Consequences
- Existing users must rename `Config\Config.env` to `Config\Config.ini`. The migration
  step is a single file rename. The old name is retained in `.gitignore` so it does not
  surface as an untracked file during the transition.
- Windows now opens the file directly in Notepad (or the default `.ini` editor) when
  double-clicked, without a file-association prompt.
- VS Code and Notepad++ apply `.ini` syntax highlighting automatically.
- Supersedes [ADR-008](ADR-008-env-config-format.md) which documented the original `.env`
  choice.
