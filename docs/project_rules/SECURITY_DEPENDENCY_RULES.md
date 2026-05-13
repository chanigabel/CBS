# Security / Dependency Rules

## Purpose

Document the project's dependency and input-safety expectations.

## Scope

- `pyproject.toml`
- `requirements.txt`
- `requirements-lock.txt`
- Excel input handling
- export sanitization helpers

## Main Files

- `pyproject.toml`
- `requirements.txt`
- `requirements-lock.txt`
- `src/excel_standardization/export/excel_safe.py`
- `webapp/services/upload_service.py`

## Responsibilities

- Keep dependencies explicit.
- Handle untrusted Excel input without executing it as code.
- Keep output values safe for Excel and filesystem consumers.

## Data Flow

1. Dependencies are installed from the project's declared files.
2. Excel files are read by libraries, not executed.
3. Sanitization helpers protect output workbooks from invalid values.

## Contracts

- Excel inputs are treated as data, not as trusted code.
- Formula-looking strings should not become active formulas in exported files.
- Illegal workbook characters should be stripped or normalized safely.

## What Must Never Change

- Untrusted Excel files must not be executed.
- Dependencies must remain pinned or declared clearly enough for reproducible
  installs.

## Current Behavior

- The project uses `openpyxl` and `xlrd` for Excel handling.
- Export safety helpers sanitize sheet names and cell contents.

## Known Limitations

- Excel files can still be malformed, password-protected, or otherwise unreadable.
- Dependency/version drift can affect packaging or runtime behavior.

## Tests That Should Cover It

- workbook safety tests
- illegal-value export tests
- upload validation tests
- packaging smoke tests

## Open Questions / Future Improvements

- Whether to add a security review checklist for new dependencies.
- Whether to document supported/unsupported file types more formally in one
  place.
