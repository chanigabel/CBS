# Agent: Upload

## Mission

Review and maintain upload validation, workbook acceptance, and session
creation.

## Files To Inspect First

- `docs/project_rules/UPLOAD_RULES.md`
- `docs/project_rules/WORKBOOK_LOADER_RULES.md`
- `webapp/services/upload_service.py`
- `webapp/api/upload.py`
- `tests/webapp/test_upload_service.py`
- `tests/webapp/test_api_upload.py`
- `tests/test_xls_legacy_support.py`

## Rules To Follow

- Keep source files untouched.
- Keep the loader policy aligned with workbook/session services.
- Treat `webapp/services/workbook_loader.py` as the canonical workbook
  dispatch path.

## What The Agent May Change

- Validation, error messages, tests, and docs when requested.

## What The Agent Must Not Change

- Business rules for the standardization engines.
- Corrected-only export policy.

## Required Tests

- Accepted `.xlsx`, `.xlsm`, `.xls` uploads.
- Invalid extension rejection.
- Corrupt workbook rejection.
- Error-path handling for legacy `.xls`.

## Regression Checklist

- Upload creates a session and sheet list.
- Original file is stored separately from the working copy.
- Invalid workbooks do not leave behind a broken session.

## Expected Output Format

- Validation findings
- loader findings
- error-path findings
- tests run and missing

## Safety Constraints

- Do not modify the uploaded source file.
- Do not rely on UI visibility when deciding loadability.
