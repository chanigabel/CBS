# Agent Gender

## 1. Mission

Review and maintain gender normalization to export codes `1` and `2`.

## 2. Files To Inspect First

- `docs/standardization_rules/GENDER_RULES.md`
- `src/excel_standardization/engines/gender_engine.py`
- `src/excel_standardization/processing/gender_standardization.py`
- `src/excel_standardization/validation/institution_report_validator.py`
- `tests/test_gender_engine.py`

## 3. Rules Documents To Follow

- `docs/standardization_rules/GENDER_RULES.md`
- `docs/standardization_rules/PIPELINE_RULES.md`
- `docs/standardization_rules/EXPORT_RULES.md`

## 4. What The Agent May Change

- Gender patterns, pipeline status handling, validation, tests, and docs when the
  requested change is explicit.

## 5. What The Agent Must Not Change

- Corrected code values `1` and `2` without approval.
- Invalid-value behavior that prevents raw invalid values from being copied into
  `gender_corrected`.

## 6. Required Safety Constraints

- Female patterns are checked before male patterns.
- Invalid non-empty values must produce empty corrected value and visible status.
- Original `gender` remains immutable.

## 7. Required Tests Before/After Changes

- `pytest tests/test_gender_engine.py`
- `pytest tests/test_institution_report_validator.py` if validation changes

## 8. Expected Output Format

List mappings changed or reviewed, invalid cases, tests run, and approval gaps.

## 9. Review Checklist

- Substring matching implications are understood.
- Empty, `None`, and whitespace-only pipeline behavior is preserved or approved.
- UI status placement is not broken.

## 10. Regression Checklist

- Numeric `1` and `2`.
- English male/female values.
- Hebrew male/female values.
- Invalid numeric values such as `8`.
- Random text values.
