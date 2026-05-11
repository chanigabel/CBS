# Agent Identifier

## 1. Mission

Review and maintain Israeli ID and passport normalization as paired identifier
logic.

## 2. Files To Inspect First

- `docs/standardization_rules/IDENTIFIER_RULES.md`
- `src/excel_standardization/engines/identifier_engine.py`
- `src/excel_standardization/processing/identifier_standardization.py`
- `src/excel_standardization/validation/institution_report_validator.py`
- `webapp/services/export_schema.py`
- `tests/test_identifier_engine.py`

## 3. Rules Documents To Follow

- `docs/standardization_rules/IDENTIFIER_RULES.md`
- `docs/standardization_rules/EXPORT_RULES.md`
- `docs/standardization_rules/INSTITUTION_RULES.md`

## 4. What The Agent May Change

- Identifier engine/pipeline code, focused tests, and docs when requested.

## 5. What The Agent Must Not Change

- Pairwise processing of ID and passport into separate single-field passes.
- Original ID/passport values.
- Passport overwrite rules without approval.

## 6. Required Safety Constraints

- Hyphen-only ID cleanup remains narrow unless approved.
- Passport-like ID values must not overwrite an existing passport value.
- Empty corrected IDs must not be counted as duplicates.

## 7. Required Tests Before/After Changes

- `pytest tests/test_identifier_engine.py`
- `pytest tests/test_export_engine_dataset.py`
- `pytest tests/test_institution_report_validator.py`

## 8. Expected Output Format

Report ID/passport routing behavior, corrected-field impact, status impact, and
tests run.

## 9. Review Checklist

- `9999` sentinel behavior.
- Too-short and too-long ID routing.
- Non-digit ID routing.
- Checksum validation.
- Passport cleanup.

## 10. Regression Checklist

- Valid checksum ID.
- Invalid checksum ID.
- All-zero ID.
- All-identical ID.
- Existing passport with moved ID.
