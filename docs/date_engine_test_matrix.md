# DateEngine Test Matrix

## Verification Commands

```powershell
pytest
python -m compileall src\excel_standardization webapp\services tests\test_date_engine_corrected_flow.py
```

Latest result:

- `920 passed, 4 skipped`
- compileall successful

## New Corrected-Flow Tests

File: `tests/test_date_engine_corrected_flow.py`

| Area | Test | Coverage |
| --- | --- | --- |
| Partial split | `test_partial_split_preserves_year_month_and_marks_missing_day` | Preserves valid year/month, blanks missing day, writes explicit status |
| Partial split | `test_partial_split_preserves_month_day_and_marks_missing_year` | Preserves month/day, blanks missing year, writes explicit status |
| Shared corrected policy | `test_single_invalid_date_uses_same_export_safe_component_policy` | Invalid single dates use corrected-field builder |
| Two-digit metadata | `test_two_digit_year_metadata_is_set_for_numeric_and_separated_paths` | Numeric and separated paths set auto-completion metadata |
| Determinism | `test_reference_year_makes_two_digit_expansion_deterministic` | Reference date controls century expansion |
| Metadata | `test_pipeline_records_processing_date_metadata` | Captured processing date/year are written to dataset metadata |
| Majority correction | `test_majority_correction_includes_single_numeric_and_separated_dates` | Single-date numeric/separated inputs participate in majority correction |
| Export contract | `test_dataset_export_uses_corrected_dates_only_without_original_fallback` | Dataset export does not resurrect original date values |

## Existing Test Coverage Preserved

### DateEngine Unit Tests

File: `tests/test_date_engine.py`

Covers:

- two-digit year expansion
- split date parsing
- numeric date parsing
- separated date parsing
- date/date-time object parsing
- business validation
- entry cutoff
- age calculation

### Date Orchestration Tests

File: `tests/test_parse_date_orchestration.py`

Covers:

- legacy `parse_date` routing
- split/main compatibility
- business validation application

### Pipeline Tests

File: `tests/test_normalization_pipeline.py`

Covers:

- split date normalization
- single date normalization
- invalid component blanking
- plain date columns
- majority correction
- no mutation of original values

### Plain Date Column Tests

File: `tests/test_plain_date_columns.py`

Covers:

- plain `birth_date`
- plain `entry_date`
- mixed split/plain configurations
- UI display column behavior
- entry-before-birth status

### Export Tests

Files:

- `tests/test_export_engine_dataset.py`
- `tests/webapp/test_export_service.py`

Covers:

- Dataset export behavior
- Web export behavior
- corrected field export paths

## Requirement Coverage

| Requirement | Covered |
| --- | --- |
| valid split dates | existing DateEngine and pipeline tests |
| invalid split dates | existing pipeline tests |
| partial split dates | new corrected-flow tests |
| non-numeric split dates | existing DateEngine and pipeline tests |
| DD/MM single dates | existing plain date tests |
| MM/DD single dates | existing DateEngine tests |
| ISO dates | existing pipeline tests |
| month-name dates | DateEngine parser coverage path retained |
| numeric compact dates | existing and new tests |
| Excel serials | existing DateEngine behavior retained; future cell-format metadata recommended |
| all parser path two-digit metadata | new corrected-flow tests cover numeric/separated; split covered by existing tests |
| deterministic expansion | new corrected-flow tests |
| processing date metadata | new corrected-flow tests |
| future dates | existing DateEngine and pipeline tests |
| age > 100 | existing DateEngine tests |
| year < 1906 | existing DateEngine tests |
| entry cutoff | existing DateEngine/plain date tests |
| entry before birth | existing plain date tests |
| majority correction split/single | existing and new tests |
| corrected-only export | new corrected-flow export test |

## Suggested Additional Tests

These are useful follow-ups if the project adds source cell-format metadata:

- integer `2020` as plain year should not be serial
- integer `36526` with date-format metadata should be serial
- integer `36526` without date-format metadata should follow configured policy
- Web export test for invalid corrected date blanking
- Institution validation test that date status is not overwritten
