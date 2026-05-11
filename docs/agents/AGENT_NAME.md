# Agent Name

## 1. Mission

Review and maintain name standardization, including text cleanup and last-name
removal from first/father name fields.

## 2. Files To Inspect First

- `docs/standardization_rules/NAME_RULES.md`
- `docs/standardization_rules/TEXT_CLEANUP_RULES.md`
- `src/excel_standardization/engines/name_engine.py`
- `src/excel_standardization/engines/text_processor.py`
- `src/excel_standardization/processing/name_standardization.py`
- `src/excel_standardization/processing/standardization_pipeline.py`
- `tests/test_name_engine.py`

## 3. Rules Documents To Follow

- `docs/standardization_rules/NAME_RULES.md`
- `docs/standardization_rules/TEXT_CLEANUP_RULES.md`
- `docs/standardization_rules/PIPELINE_RULES.md`

## 4. What The Agent May Change

- Name cleanup code, name pipeline helpers, and focused tests when requested.
- Documentation for current behavior or approved changes.

## 5. What The Agent Must Not Change

- Runtime behavior based on guessed linguistic rules.
- Original source fields.
- Export mappings for name fields without explicit approval.

## 6. Required Safety Constraints

- `first_name`, `last_name`, and `father_name` remain immutable.
- Corrections are written only to `*_corrected`.
- Pattern-based last-name removal must remain deterministic and test-covered.

## 7. Required Tests Before/After Changes

- `pytest tests/test_name_engine.py`
- `pytest tests/test_normalization_pipeline.py`
- `pytest tests/test_institution_report_validator.py` when required-name validation is touched

## 8. Expected Output Format

Summarize changed behavior, examples affected, tests run, and any remaining
ambiguity requiring approval.

## 9. Review Checklist

- Text cleanup order did not change accidentally.
- Hebrew-tie language dominance is preserved unless explicitly approved.
- Stage B last-name removal runs only when Stage A did not change the value.
- Single-word first names are protected from positional fallback.

## 10. Regression Checklist

- Hyphen/backslash separation.
- Parenthesized acronym removal.
- Hebrew/English title removal.
- First-name pattern detection.
- Father-name pattern detection.
