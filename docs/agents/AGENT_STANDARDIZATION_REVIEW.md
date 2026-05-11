# Agent Standardization Review

## 1. Mission

Perform cross-engine reviews for correctness, regressions, missing tests, and
documentation drift across the standardization system.

## 2. Files To Inspect First

- `docs/standardization_rules/README.md`
- all files in `docs/standardization_rules/`
- `DATE_RULES.md`
- `src/excel_standardization/engines/`
- `src/excel_standardization/processing/`
- `src/excel_standardization/export/`
- `src/excel_standardization/validation/`
- `webapp/services/`
- relevant tests in `tests/`

## 3. Rules Documents To Follow

All standardization rules documents. Treat conflicts as findings, not as
opportunities to invent rules.

## 4. What The Agent May Change

- Documentation and tests when requested.
- Runtime code only when the task explicitly asks for implementation changes.

## 5. What The Agent Must Not Change

- Business logic during documentation-only tasks.
- Existing user edits.
- Original-value immutability or corrected-field contracts.

## 6. Required Safety Constraints

- Clearly separate approved rule, current behavior, needs approval, and
  potential issue.
- Do not document guessed behavior as fact.
- Do not hide conflicts between code, tests, docs, UI, and export.

## 7. Required Tests Before/After Changes

Run the narrowest relevant tests for changed areas. For broad runtime changes,
run:

- `pytest tests/test_name_engine.py tests/test_gender_engine.py tests/test_identifier_engine.py`
- `pytest tests/test_date_engine.py tests/test_date_engine_corrected_flow.py`
- `pytest tests/test_institution_report_validator.py`
- relevant `tests/webapp/` tests

Documentation-only changes do not require runtime tests, but links and filenames
should be checked.

## 8. Expected Output Format

Findings first, ordered by severity:

- file/line reference
- behavior observed
- expected rule or unclear rule
- recommended next action

Then include tests run, docs updated, and approval questions.

## 9. Review Checklist

- Original fields immutable.
- Corrected fields explicit.
- Status fields visible where required.
- Export mappings explicit.
- UI grid does not leak internal keys.
- API/session flow preserves workbook data.
- Tests cover current behavior.

## 10. Regression Checklist

- Engine outputs.
- Pipeline metadata.
- Validation statuses.
- UI display columns.
- Export workbook schema.
- Date corrected-only export.
