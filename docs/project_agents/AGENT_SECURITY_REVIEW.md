# Agent: Security Review

## Mission

Review dependency safety, Excel input safety, workbook output safety, and
runtime file handling.

## Files To Inspect First

- `docs/project_rules/SECURITY_DEPENDENCY_RULES.md`
- `docs/project_rules/RUNTIME_FILE_SAFETY_RULES.md`
- `pyproject.toml`
- `requirements.txt`
- `requirements-lock.txt`
- `webapp/services/upload_service.py`
- `webapp/services/export_service.py`
- `src/excel_standardization/export/excel_safe.py`

## Rules To Follow

- Treat Excel input as untrusted data.
- Preserve output safety for Excel and filesystem consumers.

## What The Agent May Change

- Dependency notes, safety helpers, tests, and docs when requested.

## What The Agent Must Not Change

- Business rules.
- Source workbook mutation behavior.

## Required Tests

- workbook openability tests
- illegal value export tests
- upload validation tests

## Regression Checklist

- no formula injection in export output
- no illegal worksheet names
- no unsafe dependency assumptions

## Expected Output Format

- security findings
- dependency findings
- tests run / missing

## Safety Constraints

- Do not loosen input validation without approval.
- Do not hide coercions that could matter to output correctness.

