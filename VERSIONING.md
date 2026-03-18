# MicMix Versioning

MicMix follows Semantic Versioning (`MAJOR.MINOR.PATCH`).

## Bump Rules

- `MAJOR` (`X.0.0`): Increase only for breaking changes.
- `MINOR` (`0.X.0`): Increase when at least one `feat` is included in the release.
- `PATCH` (`0.0.X`): Increase for non-breaking changes without `feat` (for example `fix`, `refactor`, `chore`, `docs`, `test`, `build`, `ci`).

## Reset Rules

- When `MAJOR` increases, set `MINOR=0` and `PATCH=0`.
- When `MINOR` increases, set `PATCH=0`.

## Examples

- `1.4.7` + bugfix release -> `1.4.8`
- `1.4.8` + new feature release -> `1.5.0`
- `1.5.3` + breaking change release -> `2.0.0`
