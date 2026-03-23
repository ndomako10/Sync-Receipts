# Commit Reference

Quick guide for making commits. See [COMMIT-SCOPES.md](commit-scopes.md) for scope details.

## Format

```
type(scope): short description
```

## Commit Types

| Type | When |
|------|------|
| `feat` | New feature |
| `fix` | Bug fix |
| `docs` | Documentation only |
| `test` | Test additions or changes |
| `refactor` | Code change with no behaviour change |
| `chore` | Tooling, gitignore, config |
| `ci` | Hooks, CI pipeline |
| `perf` | Performance improvement |
| `style` | Formatting only |
| `build` | Build system changes |

## Rules

- **One commit per `type(scope)`** -- do not batch unrelated changes
- **Commit as work progresses**, not all at the end
- **Stage specific files** rather than `git add -A` to keep commits focused
- **Link to issues** in the commit body (e.g., `closes #29` or `refs #25`)
- **Propose the message and wait for approval** before committing
- **Never include Co-Authored-By trailer**
- **Never bypass pre-commit hooks** with `--no-verify`

## Versioning

This project uses [Semantic Versioning](https://semver.org): `MAJOR.MINOR.PATCH`.

- **PATCH** -- bug fix; no new parameters or changed behaviour
- **MINOR** -- new feature or parameter; backwards compatible
- **MAJOR** -- breaking change (renamed parameter, removed behaviour, changed interface)

## Changelog Workflow

1. Write a hand-crafted entry at the top of `CHANGELOG.md` **before tagging**
2. Use user-facing language grouped under sections:
   - `### Added` -- new features
   - `### Changed (breaking)` -- breaking changes
   - `### Fixed` -- bug fixes
   - `### Removed` -- deprecated features
3. The release workflow automatically publishes this entry as the GitHub Release body

**Version is tracked in git tags only** (e.g. `v1.0.0`); the script carries no version number.
