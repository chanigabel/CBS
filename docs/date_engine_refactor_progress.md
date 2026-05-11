# DateEngine Refactor Progress

## Summary

Implemented the corrected DateEngine architecture for the active Web/Dataset flow. The refactor keeps public compatibility methods in place while adding structured `DateInput`, extended `DateParseResult` metadata, deterministic reference-date handling, unified corrected-field writing, corrected partial split behavior, broader majority-correction eligibility, and aligned export behavior.

## Files Changed

- `src/excel_standardization/data_types.py`
- `src/excel_standardization/engines/date_engine.py`
- `src/excel_standardization/processing/date_standardization.py`
- `src/excel_standardization/processing/standardization_pipeline.py`
- `src/excel_standardization/export/export_engine.py`
- `webapp/services/standardization_service.py`
- `tests/test_date_engine_corrected_flow.py`
- `docs/date_engine_refactor_progress.md`
- `docs/date_engine_migration_notes.md`
- `docs/date_engine_test_matrix.md`

## Logic Changed

### Structured Date Input

Added `DateInput` as the structured internal model for DateEngine calls:

- `source_kind`
- `field_type`
- `raw_value`
- `raw_year`
- `raw_month`
- `raw_day`
- `pattern`
- `reference_date`

Legacy DateEngine methods still work.

### Extended DateParseResult

Extended `DateParseResult` with:

- `severity`
- `status_code`
- `missing_components`
- `invalid_components`
- `original_year_value`
- `original_year_digits`
- `reference_year`
- `year_was_defaulted`
- `is_calendar_valid`
- `is_business_valid`
- `source_kind`

Existing constructor usage remains compatible.

### Partial Split Dates

Any split field in the active row adapter now routes to split mode. Numeric partial split dates preserve valid components and write explicit missing-component statuses.

Examples:

- `2010 + 05 + ""` writes year/month corrected and blank day with `חסר יום`.
- `"" + 05 + 12` writes month/day corrected and blank year with `חסר שנה`.

### Shared Corrected-Field Builder

`date_corrected_components()` now uses structured result metadata and is used for both split and single date paths.

### Processing Date / Reference Date

`StandardizationPipeline` and `StandardizationService` now capture a processing date once per standardization run and provide it to DateEngine. The default is `date.today()`, and the same captured date is used for:

- two-digit year expansion
- two-part separated date defaulting
- future checks
- entry-date cutoff
- age calculation

The normalized dataset records:

- `metadata["processing_date"]`
- `metadata["processing_year"]`

### Parser Metadata Stabilization

All two-digit year parser paths now mark auto-completion metadata:

- split parsing
- numeric compact parsing
- separated string parsing
- mixed month-name parsing

### Majority Correction

Majority correction now benefits from auto-completion metadata from single-date numeric and separated parser paths, not just split paths. Changed rows are revalidated and corrected fields are rebuilt through the shared corrected-field policy.

### Export Contract

Dataset `ExportEngine` date mapping now matches Web export behavior and reads corrected date fields only. It no longer falls back from corrected date fields to original date fields.

## Risks

- DateEngine now distinguishes calendar/business validity internally while preserving `is_valid` compatibility.
- The runtime year remains business-significant; the refactor makes it explicit and consistent rather than removing it.
- Integer values between 1900 and 2100 are no longer treated as Excel serials by DateEngine without source metadata. This avoids plain-year misinterpretation but could affect files that relied on year-like serial behavior.
- Majority correction may affect more rows because single-date two-digit inputs now carry metadata.
- Partial split dates now produce explicit missing-component statuses instead of behaving like empty dates.

## Compatibility Notes

- Public DateEngine wrappers remain available:
  - `parse_date`
  - `parse_from_split_columns`
  - `parse_from_main_value`
  - `parse_date_value`
  - `parse_numeric_date_string`
  - `parse_separated_date_string`
  - `expand_two_digit_year`
- Existing corrected field names are preserved.
- Existing date status field names are preserved.
- Existing export schema field names are preserved.
- Original raw row fields are preserved.

## Tests Added

Added `tests/test_date_engine_corrected_flow.py` covering:

- partial split component preservation
- shared corrected-field policy for invalid single dates
- two-digit year metadata for numeric and separated parser paths
- deterministic reference-year expansion
- majority correction across single numeric/separated dates
- Dataset export corrected-only date behavior

## Verification

Commands run:

```powershell
pytest
python -m compileall src\excel_standardization webapp\services tests\test_date_engine_corrected_flow.py
```

Results:

- `920 passed, 4 skipped`
- compileall completed successfully

## Remaining Work

- Consider adding source cell-format metadata so Excel serial handling can distinguish plain integer years from true serial dates.
- Consider replacing Hebrew display-text comparisons with stable status codes in UI-facing tests over time.
- Decide whether majority correction should remain one-way only or become bidirectional under an approved business rule.
