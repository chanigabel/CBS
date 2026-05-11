# Agent Pipeline

## 1. Mission

Review and maintain the end-to-end standardization pipeline and web session flow.

## 2. Files To Inspect First

- `docs/standardization_rules/PIPELINE_RULES.md`
- `src/excel_standardization/processing/standardization_pipeline.py`
- `src/excel_standardization/processing/*_standardization.py`
- `src/excel_standardization/workbook_json_flow.py`
- `webapp/services/standardization_service.py`
- `webapp/services/workbook_service.py`
- `webapp/services/processing_report_service.py`

## 3. Rules Documents To Follow

- `docs/standardization_rules/PIPELINE_RULES.md`
- all engine-specific rules touched by the change
- `docs/standardization_rules/EXPORT_RULES.md`

## 4. What The Agent May Change

- Pipeline orchestration, service flow, metadata/statistics, tests, and docs
  when requested.

## 5. What The Agent Must Not Change

- Engine rules indirectly while refactoring orchestration.
- Original source fields.
- UI visibility of internal helper keys unless approved.

## 6. Required Safety Constraints

- Engine failures must remain isolated where current behavior expects recovery.
- At least one normalized sheet must succeed for partial success.
- Manual edit replay behavior must be considered before changing order.

## 7. Required Tests Before/After Changes

- `pytest tests/test_normalization_pipeline.py`
- `pytest tests/webapp/test_normalization_service.py`
- `pytest tests/webapp/test_workbook_service.py`
- relevant engine tests

## 8. Expected Output Format

Describe stage order impact, metadata/status impact, UI/export impact, and tests
run.

## 9. Review Checklist

- Engine enable flags.
- Dataset-level pattern detection.
- Validation timing.
- Metadata statistics.
- Internal helper key stripping.

## 10. Regression Checklist

- Single-sheet standardize.
- Full-workbook standardize.
- Partial sheet failure.
- Manual edit replay.
- Processing report stage updates.
