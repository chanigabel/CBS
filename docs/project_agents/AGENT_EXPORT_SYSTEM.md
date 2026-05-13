# Agent: Export System

## Mission

Review and maintain all workbook export paths and their safety helpers.

## Files To Inspect First

- `docs/project_rules/EXPORT_SYSTEM_RULES.md`
- `webapp/services/export_service.py`
- `webapp/services/export_writer.py`
- `webapp/services/export_rows.py`
- `webapp/services/export_schema.py`
- `src/excel_standardization/export/export_engine.py`
- `src/excel_standardization/export/excel_safe.py`

## Rules To Follow

- Export must remain corrected-only for standardized columns.
- Exported files must be valid `.xlsx` workbooks.

## What The Agent May Change

- Export assembly, safety helpers, tests, and docs when requested.

## What The Agent Must Not Change

- Standardization engine rules.
- Source workbook data.

## Required Tests

- active export tests
- compatibility export tests
- workbook openability tests
- sanitized sheet-name collision tests

## Regression Checklist

- corrected-field mapping
- safe sheet titles
- safe cell values
- filename sanitization

## Expected Output Format

- export schema findings
- workbook validity findings
- tests run / missing

## Safety Constraints

- Do not depend on UI visibility.
- Do not mutate source rows in place unless the current contract already
  requires it and it is documented.

