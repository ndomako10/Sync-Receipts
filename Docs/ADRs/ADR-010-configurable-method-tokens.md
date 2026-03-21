# ADR-010: Configurable Payment Method Tokens

## Status
Accepted

## Context

Payment method tokens (`Card`, `Cash`, `Checking`, `Savings`, etc.) were hardcoded in two
places:

1. The regex in `ConvertFrom-ReceiptFileName` -- determines which tokens are accepted at
   parse time.
2. The dropdown validation list in `New-AccountsTemplate.ps1` -- determines which tokens
   appear as choices in `Config/Templates/Accounts.template.xlsx`.

Having two independent hardcoded lists creates a drift risk: a token added to the dropdown
but not the regex (or vice versa) produces an inconsistent user experience. It also means
adding a new method requires a code change and a new release.

An audit of the parser revealed that there is no per-method branching logic. All non-Cash
methods are treated identically: the token is captured and written to the sheet. The only
special case is `Cash`, which is exempt from the Last 4 requirement. This absence of
branching logic means a user-defined token carries no risk of corrupting account lookup or
flag behaviour -- the parser simply accepts it and passes it through.

`Categories.json` establishes a precedent for user-configurable data that drives both
runtime behaviour (the category dropdown) and template generation, following the pattern:
committed template -> gitignored local copy created by `Initialize-SyncReceipts.ps1`.

## Decision

Extract the non-Cash method tokens into `Config/Methods.json`. Both
`ConvertFrom-ReceiptFileName` and `New-AccountsTemplate.ps1` read from this file at
runtime. `Cash` remains hardcoded in the parser as the sole token with distinct behaviour
(Last 4 exemption).

### File format

```json
[
  "Card",
  "Check",
  "Checking",
  "Savings",
  "Transfer",
  "Wire"
]
```

`Cash` is intentionally absent -- it is always valid and always exempt from Last 4
validation regardless of the contents of this file.

### Behaviour

- `ConvertFrom-ReceiptFileName` reads `Methods.json` once per invocation and builds the
  regex dynamically. If the file is absent or unreadable, the function falls back to the
  built-in default token set and logs a warning.
- `New-AccountsTemplate.ps1` reads `Methods.json` to populate the Method dropdown,
  prepending `Cash` so it always appears first in the list.
- `Initialize-SyncReceipts.ps1` copies `Config/Templates/Methods.template.json` to
  `Config/Methods.json` during setup (skipped if already present), matching the
  `Categories.json` pattern.
- `Config/Methods.json` is gitignored; `Config/Templates/Methods.template.json` is committed.

### Constraints

- Token values must be single words (no spaces, no special regex characters). The parser
  validates each token before building the regex and skips invalid entries with a warning.
- `Cash` must not appear in `Methods.json` -- its special-case handling is hardcoded and
  adding it to the file would have no effect.

## Alternatives considered

- **Feed dropdown only, keep parser hardcoded** -- rejected because the dropdown and parser
  would drift if a user adds a custom token. A token that appears in the dropdown but is
  rejected by the parser is a confusing failure mode.
- **Keep both hardcoded** -- simple, but requires a code change and release for any new
  method. Unjustified given that no per-method branching logic exists.
- **Single shared constant in the script** -- eliminates drift between the two hardcoded
  lists but does not give users extensibility. A partial improvement rejected in favour of
  the full solution.

## Flag behaviour

A receipt filename containing an unrecognised method token fails the parser regex and is
flagged as `"Could not parse filename"` -- the same result as any other structural parse
failure. This is consistent with the current behaviour for hardcoded tokens and loses no
existing flag specificity.

A more descriptive `"Unrecognised method"` flag is not currently implemented. Adding it
would require a two-pass parse: a broad regex that accepts any word as the method token,
followed by a check of the token against the loaded Methods list. See issue #63.

## Consequences

- `ConvertFrom-ReceiptFileName` gains a file dependency and is no longer a pure function.
  Unit tests must supply a Methods list explicitly (via a new `-Methods` parameter) to
  remain file-independent.
- A misconfigured `Methods.json` (e.g. a token containing spaces or regex metacharacters)
  can cause parse failures. The parser guards against this by validating tokens before
  building the regex.
- Adding a new built-in method in future requires updating `Config/Templates/Methods.template.json`
  and regenerating `Config/Templates/Accounts.template.xlsx` via `New-AccountsTemplate.ps1`; no parser code
  change is needed.
- `Initialize-SyncReceipts.ps1` gains a new copy step for `Methods.json`, consistent with
  the existing `Categories.json` and `Accounts.xlsx` steps.
