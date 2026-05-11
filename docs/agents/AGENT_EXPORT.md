# Agent Export

## 1. Mission

Review and maintain the workbook export contract.

## 2. Files To Inspect First

- `docs/standardization_rules/EXPORT_RULES.md`
- `webapp/services/export_service.py`
- `webapp/services/export_writer.py`
- `webapp/services/export_rows.py`
- `webapp/services/export_schema.py`
- `webapp/services/export_validation.py`
- `src/excel_standardization/export/export_engine.py`
- `tests/webapp/test_export_service.py`

## 3. Rules Documents To Follow

- `docs/standardization_rules/EXPORT_RULES.md`
- `docs/standardization_rules/PIPELINE_RULES.md`
- engine-specific rules for mapped fields

## 4. What The Agent May Change

- Export schema, row filtering, export writer, export tests, and docs when
  requested.

## 5. What The Agent Must Not Change

- Standardization engine business logic.
- Source workbook data.
- Corrected-field preference without approval.

## 6. Required Safety Constraints

- Export should not become a normalization engine.
- Active web export maps explicitly through `EXPORT_MAPPING`.
- Final workbook status-column behavior must be deliberate and documented.

## 7. Required Tests Before/After Changes

- `pytest tests/webapp/test_export_service.py`
- `pytest tests/webapp/test_api_export.py`
- `pytest tests/test_export_engine_dataset.py`
- date corrected-flow export tests when dates are touched

## 8. Expected Output Format

List schema changes, row filtering changes, generated workbook impact, and tests
run.

## 9. Review Checklist

- Headers and order.
- Sheet canonicalization.
- MosadID and SugMosad injection.
- Corrected-field mapping.
- Empty/numeric helper row filtering.

## 10. Regression Checklist

- Dayarim export.
- Meshkey export with Dira.
- Anashey export with Dira.
- Unknown sheets.
- Moved passport values.
