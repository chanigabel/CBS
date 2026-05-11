# Agent Date

## 1. Mission

Review and maintain date standardization behavior without changing the approved
date contract accidentally.

## 2. Files To Inspect First

- `DATE_RULES.md`
- `docs/standardization_rules/DATE_RULES_REFERENCE.md`
- `src/excel_standardization/engines/date_engine.py`
- `src/excel_standardization/processing/date_standardization.py`
- `src/excel_standardization/processing/standardization_pipeline.py`
- `webapp/services/workbook_service.py`
- `webapp/services/export_schema.py`
- `webapp/services/export_writer.py`

## 3. Rules Documents To Follow

- `DATE_RULES.md`
- `docs/standardization_rules/PIPELINE_RULES.md`
- `docs/standardization_rules/EXPORT_RULES.md`

## 4. What The Agent May Change

- Date parsing or pipeline code only when the requested task explicitly requires it.
- Date tests and date docs.
- Status wording only with approval, because statuses are user-visible Hebrew text.

## 5. What The Agent Must Not Change

- Original source date fields.
- Export date mapping from corrected components without explicit approval.
- Business date rules by inference.

## 6. Required Safety Constraints

- Original values remain immutable.
- Corrections go only into corrected fields.
- Suspicious, recovered, ambiguous, invalid, or special-source dates must remain visible through status fields.
- Internal helper fields must not leak into the UI grid.

## 7. Required Tests Before/After Changes

- `pytest tests/test_date_engine.py`
- `pytest tests/test_date_engine_corrected_flow.py`
- `pytest tests/test_date_conservative_parsing.py`
- `pytest tests/test_plain_date_columns.py`
- relevant web tests when UI/export behavior changes

## 8. Expected Output Format

Report findings first, with file and line references. Then list behavior impact,
tests run, and any open approval questions.

## 9. Review Checklist

- Corrected date components are used for UI and export.
- Invalid components are blanked according to `date_corrected_components`.
- Date statuses survive later pipeline stages.
- Single-date and split-date inputs are both covered.

## 10. Regression Checklist

- Compact numeric dates.
- Excel serial dates.
- Missing components.
- Impossible dates.
- Entry-before-birth warning.
- Birth-year majority correction.
