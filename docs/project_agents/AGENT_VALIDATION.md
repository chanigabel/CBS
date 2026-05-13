# Agent: Validation

## Mission

Review and maintain institution / Mosad validation over normalized workbook
rows.

## Files To Inspect First

- `docs/project_rules/VALIDATION_RULES.md`
- `docs/project_rules/STANDARDIZATION_PIPELINE_RULES.md`
- `src/excel_standardization/validation/institution_report_validator.py`
- `webapp/services/standardization_service.py`
- `tests/test_institution_report_validator.py`

## Rules To Follow

- Prefer corrected values when validating normalized output.
- Keep workbook-wide checks workbook-wide.
- Treat `src/excel_standardization/normalized_row_contract.py` as the shared
  source for validation-source selection.

## What The Agent May Change

- Validation helpers, messages, tests, and docs when requested.

## What The Agent Must Not Change

- Standardization engine business rules.
- Source workbook values.

## Required Tests

- row-level validation tests
- workbook-level duplicate tests
- mixed corrected/original tests
- type-edge-case tests

## Regression Checklist

- validation still writes `_validation_status`
- corrected fields remain visible
- workbook duplicate checks still work

## Expected Output Format

- validation findings
- message/constant findings
- tests run / missing

## Safety Constraints

- Keep validation side effects deliberate and minimal.
- Do not suppress actual validation failures.
