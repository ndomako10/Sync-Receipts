Review CHANGELOG.md for completeness, accuracy, and format compliance.

Read the following files:
- CHANGELOG.md
- CONTRIBUTING.md (versioning and changelog section only)

Also run:
- `git tag --sort=-version:refname | head -10` to list recent tags
- `LATEST=$(git tag --sort=-version:refname | head -1); PREV=$(git tag --sort=-version:refname | head -2 | tail -1); echo "Latest: $LATEST  Prev: $PREV"`
- `git log --oneline $LATEST..HEAD` to list commits since the last tag
- `git log --oneline $PREV..$LATEST` to list commits in the most recent release
- `gh issue list --state open --label changelog` to check for known outstanding changelog gaps

Check each area and flag any issues:

1. **Format compliance** -- Does CHANGELOG.md follow Keep a Changelog conventions?
   - Top entry is `## [Unreleased]` or a versioned heading `## [x.y.z] - YYYY-MM-DD`
   - Changes grouped under `### Added`, `### Changed`, `### Fixed`, `### Removed` (only sections with content are required)
   - No machine-generated or commit-message-style entries -- language should be user-facing

2. **Most recent release entry** -- Compare the commits in the most recent release (from git log above) against the changelog entry for that version:
   - Are all user-facing changes represented?
   - Are any entries present that don't correspond to a real change (stale or inaccurate)?
   - Do breaking changes appear under `### Changed` with a clear description of the impact?

3. **Unreleased section** -- Compare commits since the last tag against the `## [Unreleased]` section (if present):
   - Are there unreleased commits with user-facing changes that are not yet documented?
   - Propose entries for any missing changes, grouped correctly by type

4. **Version consistency** -- Does the version in the most recent changelog entry match the most recent git tag? Flag any mismatch.

5. **Semantic versioning** -- Are version bumps appropriate for the changes described?
   - Breaking changes (renamed parameters, removed behaviour, changed file format) warrant a MAJOR bump
   - New features or parameters warrant a MINOR bump
   - Bug fixes, docs, and tests warrant a PATCH bump

## Output format

For each finding: `SEVERITY | FILE:LINE | one-sentence problem | one-sentence fix`
Severity: CRITICAL (broken/data-loss), MAJOR (correctness gap), MINOR (polish).
If no issues in an area: `OK: <area>`
No narrative. No file contents. Findings only.
