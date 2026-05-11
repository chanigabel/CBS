# DateEngine Corrected Flow Design

Scope: active Web/Dataset runtime flow only.

This document describes the ideal corrected DateEngine architecture and runtime flow after the review findings are addressed. It is intentionally forward-looking: it describes how the flow should work after fixes, not a restatement of the current implementation.

The target behavior is:

- Raw Excel values are preserved.
- Date parsing is deterministic and source-aware.
- Split, partial split, and single-date inputs share one result model.
- Corrected fields are written consistently for all input types.
- Status fields explain DateEngine outcomes.
- Institution validation remains separate from DateEngine status.
- Export writes finalized corrected fields only.
- Two-digit year handling and majority correction are deterministic within a standardization run because the processing date is captured once and passed consistently.

## Processing Date Policy

The corrected architecture does not remove runtime-date behavior. In this business domain, the processing date is intentionally meaningful because every reporting year can have different expected values and validity rules.

The corrected policy is:

1. Capture `processing_date` once at the start of a standardization run.
2. Default `processing_date` to `date.today()`.
3. Store the captured date on the pipeline as the run reference date.
4. Pass the same date to DateEngine for every row and every parser path.
5. Use it consistently for:
   - two-digit year expansion
   - future-date validation
   - entry-date cutoff
   - two-part separated date default year
   - tests
   - logs and metadata
6. Record it in dataset metadata as:
   - `processing_date`
   - `processing_year`

This makes output deterministic for a given run and a given processing date. It does not make the system independent of the runtime year; it makes the runtime year explicit and testable.

## Corrected End-To-End Runtime Flow

### Purpose

The end-to-end flow should transform raw Excel date values into export-ready corrected components while preserving the original values and recording clear status and validation metadata.

### Correct Execution Order

1. Extract worksheet values into a `SheetDataset`.
2. Detect canonical date fields:
   - split date fields: `birth_year`, `birth_month`, `birth_day`, `entry_year`, `entry_month`, `entry_day`
   - single date fields: `birth_date`, `entry_date`
3. Detect sheet-level separated-date pattern only for ambiguous separated strings.
4. Normalize rows through `StandardizationPipeline.normalize_dataset`.
5. For each row, normalize birth date and entry date separately.
6. For each date group, build an explicit date input model:
   - source kind: `split`, `single`, or `missing`
   - raw values
   - field type: `birth` or `entry`
   - processing date/reference date
   - detected date pattern
7. DateEngine parses the input into date components.
8. DateEngine validates component structure and calendar validity.
9. DateEngine applies date business rules.
10. The adapter writes corrected fields and date status fields with one shared policy.
11. After all rows are normalized, majority correction adjusts eligible auto-completed birth years.
12. Changed rows are revalidated after majority correction.
13. Institution validation writes `_validation_ok` and `_validation_status`.
14. Export writes only finalized corrected fields.

### Inputs

Examples of raw extracted rows:

```python
{"birth_year": "24", "birth_month": "02", "birth_day": "01"}
{"birth_year": "2010", "birth_month": "05", "birth_day": ""}
{"birth_date": "01/02/24"}
{"birth_date": "010224"}
{"entry_date": 36526}
{"entry_date": datetime(2020, 9, 15)}
```

### Outputs

Rows should contain original fields plus corrected and status fields:

```python
{
    "birth_date": "01/02/24",
    "birth_year_corrected": 2024,
    "birth_month_corrected": 2,
    "birth_day_corrected": 1,
    "birth_date_status": "",
    "_validation_ok": True,
    "_validation_status": ""
}
```

### Ownership

- Extraction owns raw value preservation.
- Date adapter owns row-to-DateEngine translation.
- DateEngine owns parsing and date-specific validation.
- Pipeline owns sheet-level correction and validator ordering.
- Institution validator owns report-level validation.
- Export owns writing finalized corrected fields.

### Bugs Fixed

- Partial split dates no longer become empty dates.
- Single and split flows no longer write corrected fields differently.
- Two-digit years in all parser paths can participate in majority correction.
- Runtime-year dependency is made explicit and testable.
- Export no longer reintroduces original invalid values.

## Corrected Responsibility Split

### `webapp/services/standardization_service.py`

Purpose:

Own web runtime orchestration.

Should own:

- Loading the session record.
- Re-extracting one sheet or the full workbook.
- Building the pipeline.
- Capturing `processing_date` once at the start of the standardization run.
- Defaulting `processing_date` to `date.today()`.
- Passing that same processing date into the pipeline and DateEngine.
- Storing normalized datasets back into the session.

Should not own:

- Date parsing.
- Date correction policy.
- Date validation rules.
- Export fallback behavior.

Inputs:

```python
session_id
sheet_name
working_copy_path
```

Outputs:

Updated session dataset containing normalized sheets.

Bugs fixed:

- Keeps DateEngine deterministic within the run while preserving the business rule that the current processing year matters.
- Makes the processing date visible in metadata and tests.

### `src/excel_standardization/io_layer/excel_to_json_extractor.py`

Purpose:

Convert Excel sheets into canonical `SheetDataset` rows while preserving raw value types.

Should own:

- Reading Excel cells.
- Preserving values as `str`, `int`, `float`, `date`, or `datetime` where appropriate.
- Mapping workbook columns to canonical field names.
- Producing `SheetDataset`.

Should not own:

- Parsing date strings.
- Expanding two-digit years.
- Applying business validation.

Inputs:

Openpyxl worksheet rows and detected columns.

Outputs:

```python
SheetDataset(
    field_names=[...],
    rows=[...],
    metadata={...}
)
```

Bugs fixed:

- Keeps raw Excel values available without mixing extraction and correction responsibilities.

### `src/excel_standardization/processing/standardization_pipeline.py`

Purpose:

Coordinate dataset-level standardization.

Should own:

- Detecting sheet-level separated-date pattern.
- Calling `normalize_row`.
- Running majority correction after all rows.
- Running institution validation after majority correction.
- Writing dataset metadata and statistics.

Should not own:

- Low-level date parsing.
- Corrected component sanitization.
- Export behavior.

Inputs:

Raw `SheetDataset`.

Outputs:

Normalized `SheetDataset`.

Bugs fixed:

- Ensures majority correction happens only after all row-level DateEngine metadata exists.
- Ensures validation sees final corrected values.

### `src/excel_standardization/processing/date_standardization.py`

Purpose:

Bridge row dictionaries and DateEngine.

Should own:

- Deciding whether a row date group is split, single, or missing.
- Building explicit DateEngine input.
- Calling DateEngine.
- Writing corrected fields.
- Writing date status fields.
- Writing temporary majority-correction metadata.
- Running entry-before-birth relation check.
- Applying sheet-level majority correction.

Should not own:

- Parsing implementation details.
- Calendar rules.
- Business thresholds.
- Institution validation.
- Export writing.

Inputs:

Row dictionary and date group metadata.

Outputs:

Mutated row dictionary with corrected/status fields.

Bugs fixed:

- Gives split and single flows one corrected-field lifecycle.
- Prevents partial split values from falling through to main-value parsing.

### `src/excel_standardization/engines/date_engine.py`

Purpose:

Parse and validate date values independently of row dictionaries.

Should own:

- Source-aware parsing.
- Split date component parsing.
- Single value dispatch.
- Numeric string parsing.
- Separated string parsing.
- Month-name parsing.
- Excel serial handling.
- Two-digit year expansion using explicit reference year.
- Calendar validation.
- Date business validation.
- Entry-before-birth relation validation.

Should not own:

- Mutating dataset rows.
- Knowing export header names.
- Writing `_validation_status`.

Inputs:

Structured date input:

```python
DateInput(
    source_kind="split",
    field_type=BIRTH_DATE,
    raw_year="24",
    raw_month="02",
    raw_day="01",
    pattern=DateFormatPattern.DDMM,
    reference_date=date(2026, 5, 11)
)
```

Outputs:

Structured `DateParseResult`.

Bugs fixed:

- Eliminates parser paths that expand two-digit years without metadata.
- Removes repeated hidden `date.today()` calls from low-level parsing. The runtime date is still used intentionally, but it is captured once as `processing_date`.

### `src/excel_standardization/validation/institution_report_validator.py`

Purpose:

Validate normalized rows against report requirements.

Should own:

- Required-field checks.
- Report-specific date component checks.
- `_validation_ok`.
- `_validation_status`.

Should not own:

- Re-parsing raw date strings.
- Rewriting corrected fields.
- Export behavior.

Inputs:

Rows after DateEngine and majority correction.

Outputs:

Rows with validation fields:

```python
{
    "_validation_ok": False,
    "_validation_status": "ShnatLida missing"
}
```

Bugs fixed:

- Keeps institution validation separate from DateEngine status.
- Avoids duplicated parsing responsibility.

### `webapp/services/export_schema.py`

Purpose:

Define export header-to-field mapping.

Should own:

- Mapping official export headers to finalized corrected fields.

Should not own:

- Date validation policy.
- Date fallback policy.

Date mapping:

```python
"ShnatLida": "birth_year_corrected"
"HodeshLida": "birth_month_corrected"
"YomLida": "birth_day_corrected"
"shnatknisa": "entry_year_corrected"
"Hodeshknisa": "entry_month_corrected"
"YomKnisa": "entry_day_corrected"
```

Bugs fixed:

- Makes corrected fields the export contract.

### `webapp/services/export_writer.py`

Purpose:

Write the active Web/Dataset export workbook.

Should own:

- Creating workbook sheets.
- Writing headers.
- Reading mapped row values.
- Writing non-empty finalized corrected fields.

Should not own:

- Re-parsing dates.
- Revalidating dates.
- Falling back to original date values.
- Repairing corrected fields.

Bugs fixed:

- Export no longer hides upstream correction issues or reintroduces original invalid values.

## Corrected Parsing Flow

### Stage Purpose

Parsing should convert raw input into structured date components and metadata, without mutating rows and without deciding export behavior.

### Inputs

```python
DateInput(
    source_kind="split" | "single" | "missing",
    field_type=BIRTH_DATE | ENTRY_DATE,
    raw_value=None,
    raw_year=None,
    raw_month=None,
    raw_day=None,
    pattern=DateFormatPattern.DDMM,
    reference_date=date(2026, 5, 11)
)
```

### Outputs

```python
DateParseResult(
    source_kind="split",
    year=2024,
    month=2,
    day=1,
    is_calendar_valid=True,
    is_business_valid=True,
    severity="ok",
    status_code="ok",
    status_text="",
    year_was_auto_completed=True,
    original_year_value=24,
    original_year_digits=2,
    missing_components=[],
    invalid_components=[]
)
```

### Correct Decision Tree

```text
Date input arrives
|
|-- source_kind = split
|   |
|   |-- parse each component independently
|   |-- preserve valid parsed components
|   |-- report missing components
|   |-- expand two-digit year with metadata
|   |-- validate calendar date if complete
|   |-- apply business rules if calendar-valid
|
|-- source_kind = single
|   |
|   |-- None / empty -> empty result
|   |-- datetime/date -> direct components
|   |-- Excel serial -> from_excel with serial metadata
|   |-- month name -> month-name parser
|   |-- all digits -> compact numeric parser
|   |-- ISO-like -> ISO parser
|   |-- separated -> separated parser
|   |-- otherwise -> unrecognized format
|
|-- source_kind = missing
|   |
|   |-- no corrected components
|   |-- optional/required decision deferred to business/institution rules
```

### What Should Not Happen

- Parser should not mutate row dictionaries.
- Parser should not write export fields.
- Parser should not call `date.today()` directly for each decision. It should use the captured `processing_date`.
- Parser should not lose metadata when expanding a year.
- Parser should not treat partial split dates as single-date empty values.

### Bugs Fixed

- Fixes inconsistent parser metadata.
- Fixes majority correction blind spots.
- Fixes inconsistent runtime-year usage by making the processing date explicit and consistent.

## Corrected Handling: Split Dates

### Purpose

Support dates already separated into year, month, and day columns.

### Inputs

```python
{
    "birth_year": "24",
    "birth_month": "02",
    "birth_day": "01"
}
```

### Ownership

- `date_standardization.normalize_date_field` identifies split source.
- `DateEngine.parse_from_split_columns` parses components.

### Correct Output

```python
{
    "birth_year_corrected": 2024,
    "birth_month_corrected": 2,
    "birth_day_corrected": 1,
    "birth_date_status": "",
    "_birth_year_auto_completed": True
}
```

### Bugs Fixed

- Split date parsing remains source-aware.
- Two-digit split years participate in majority correction.

## Corrected Handling: Partial Split Dates

### Purpose

Preserve valid components while reporting missing or invalid components precisely.

### Inputs

```python
{
    "birth_year": "2010",
    "birth_month": "05",
    "birth_day": ""
}
```

### Ownership

- `date_standardization.normalize_date_field` must route any split-field presence to split mode.
- `DateEngine.parse_from_split_columns` must support partial components.

### Correct Output

```python
{
    "birth_year_corrected": 2010,
    "birth_month_corrected": 5,
    "birth_day_corrected": "",
    "birth_date_status": "missing day",
    "_birth_year_auto_completed": False
}
```

### What Should Not Happen

- Do not call single-value parsing with `main_val=None`.
- Do not discard valid year/month.
- Do not mark the result as simply empty.

### Bugs Fixed

- Fixes partial split data loss.
- Fixes inaccurate empty-cell statuses.
- Improves UI and validation diagnostics.

## Corrected Handling: Single Dates

### Purpose

Parse one-cell dates into corrected year/month/day components.

### Inputs

```python
{"birth_date": "01/02/24"}
```

### Ownership

- `date_standardization.normalize_date_field` identifies single source.
- `DateEngine.parse_date_value` dispatches by value type/content.
- Shared corrected-field builder writes outputs.

### Correct Output

```python
{
    "birth_year_corrected": 2024,
    "birth_month_corrected": 2,
    "birth_day_corrected": 1,
    "birth_date_status": "",
    "_birth_year_auto_completed": True
}
```

### What Should Not Happen

- Single flow must not write `result.year`, `result.month`, and `result.day` directly with a different safety policy.
- Single flow must not skip auto-completion metadata.

### Bugs Fixed

- Makes single and split corrected-field behavior consistent.
- Enables majority correction for single-date two-digit years.

## Corrected Handling: Numeric Strings

### Purpose

Parse compact numeric dates.

### Inputs

```python
"01022024"
"010224"
"1225"
"2020"
```

### Ownership

- `DateEngine._parse_numeric_date_string`

### Correct Behavior

```text
01022024 -> 01/02/2024
010224   -> 01/02/2024, auto-completed year
1225     -> 1/2/2025, auto-completed year
2020     -> partial year-only date
```

### Correct Output Example

Input:

```python
{"birth_date": "010224"}
```

Output:

```python
{
    "birth_year_corrected": 2024,
    "birth_month_corrected": 2,
    "birth_day_corrected": 1,
    "birth_date_status": "",
    "_birth_year_auto_completed": True
}
```

### What Should Not Happen

- Compact numeric two-digit years must not be expanded without metadata.
- Compact numeric strings should not use separated-date pattern unless explicitly designed later.

### Bugs Fixed

- Majority correction includes numeric single-date rows.

## Corrected Handling: Separated Dates

### Purpose

Parse slash/dot-separated date strings using a detected or configured date order.

### Inputs

```python
"01/02/24"
"01.02.2024"
"1/2"
```

### Ownership

- `date_standardization.detect_date_format_pattern`
- `DateEngine._parse_separated_date_string`

### Correct Behavior

For `DateFormatPattern.DDMM`:

```python
"01/02/24" -> day=1, month=2, year=2024
```

For `DateFormatPattern.MMDD`:

```python
"01/02/24" -> month=1, day=2, year=2024
```

For two-part dates:

```python
"1/2" -> use explicit reference_year, mark year_was_defaulted=True
```

### Correct Output Example

```python
{
    "birth_year_corrected": 2024,
    "birth_month_corrected": 2,
    "birth_day_corrected": 1,
    "birth_date_status": "",
    "_birth_year_auto_completed": True
}
```

### What Should Not Happen

- Do not call `date.today()` inside the parser for each value. Use the captured processing date.
- Do not silently default a missing year without metadata.
- Do not expand a two-digit year without marking auto-completion.

### Bugs Fixed

- Makes separated date parsing deterministic.
- Makes majority correction work for separated strings.

## Corrected Handling: Excel Serials

### Purpose

Handle Excel serial numbers when the source value is truly a date serial.

### Inputs

```python
36526
```

### Ownership

- `DateEngine.parse_date_value`
- Ideally supported by extraction metadata indicating cell date format.

### Correct Behavior

If the value is confirmed as an Excel date serial:

```python
36526 -> date from openpyxl.utils.datetime.from_excel
```

Correct output:

```python
{
    "entry_year_corrected": 2000,
    "entry_month_corrected": 1,
    "entry_day_corrected": 1,
    "entry_date_status": ""
}
```

### What Should Not Happen

- Plain integer years like `2020` should not automatically become Excel serial dates unless source metadata supports that interpretation.

### Bugs Fixed

- Prevents type-based misinterpretation of plain integer year-like values.

## Corrected Handling: Two-Digit Years

### Purpose

Expand shortened years deterministically and preserve metadata for majority correction.

### Inputs

```python
"24"
"010224"
"01/02/24"
```

### Ownership

- `DateEngine._expand_two_digit_year`
- All parser methods that call it

### Correct Behavior

Use explicit reference year:

```python
reference_year = 2026
```

Expansion:

```python
24 -> 2024
26 -> 2026
27 -> 1927
99 -> 1999
```

Every expansion should set:

```python
year_was_auto_completed = True
original_year_value = 24
original_year_digits = 2
reference_year = 2026
```

### What Should Not Happen

- Do not use fresh `date.today()` calls inside parsing. Use the captured processing date.
- Do not expand without marking metadata.

### Bugs Fixed

- Deterministic behavior across years.
- Complete majority-correction eligibility.

## Corrected Handling: Future Dates

### Purpose

Apply business rules after calendar-valid parsing.

### Inputs

```python
DateParseResult(year=2027, month=1, day=1, is_calendar_valid=True)
field_type = BIRTH_DATE
reference_date = date(2026, 5, 11)
```

### Ownership

- `DateEngine.validate_business_rules`

### Correct Output

```python
DateParseResult(
    year=2027,
    month=1,
    day=1,
    is_calendar_valid=True,
    is_business_valid=False,
    severity="error",
    status_code="future_birth_date",
    status_text="future birth date"
)
```

### What Should Not Happen

- Business validation should not run before calendar validation.
- Future checks should use the captured processing date, not a fresh hidden runtime-date lookup.
- Export should not decide future-date behavior.

### Bugs Fixed

- Makes future-date behavior deterministic and testable.

## Corrected Handling: Invalid Dates

### Purpose

Represent invalid dates clearly without exporting unsafe corrected values.

### Inputs

```python
{"birth_date": "31/02/2020"}
```

### Ownership

- DateEngine parses and validates.
- Date adapter writes corrected fields through one shared policy.

### Correct Output

Component-preserving policy example:

```python
{
    "birth_year_corrected": 2020,
    "birth_month_corrected": 2,
    "birth_day_corrected": "",
    "birth_date_status": "date does not exist"
}
```

Strict policy example:

```python
{
    "birth_year_corrected": "",
    "birth_month_corrected": "",
    "birth_day_corrected": "",
    "birth_date_status": "date does not exist"
}
```

The chosen policy must be shared by all input types.

### What Should Not Happen

- Single-date flow must not export `31` as a corrected day.
- Split and single invalid dates must not behave differently.

### Bugs Fixed

- Prevents invalid corrected values from reaching export.

## Corrected Validation Flow

### Purpose

Validation should be layered and explicit.

### Layers

1. Parse validation:
   - can raw values be converted?
2. Component validation:
   - missing year/month/day
   - invalid numeric ranges
3. Calendar validation:
   - impossible dates like February 31
4. Date business validation:
   - future date
   - entry cutoff
   - minimum year
   - age warning/error
5. Cross-date validation:
   - entry before birth
6. Institution validation:
   - report-required fields
   - `_validation_ok`
   - `_validation_status`

### Ownership

- DateEngine owns layers 1-5 except row mutation.
- Institution validator owns layer 6.
- Export owns none of these.

### Bugs Fixed

- Prevents duplicated or contradictory validation.
- Separates DateEngine statuses from report validation fields.

## Corrected Corrected-Field Lifecycle

### Purpose

Corrected fields should be the single source of truth for UI and export.

### Lifecycle

1. Raw fields enter the row.
2. DateEngine returns structured `DateParseResult`.
3. Shared corrected-field builder produces:
   - `year_corrected`
   - `month_corrected`
   - `day_corrected`
   - `date_status`
4. Adapter writes:
   - `birth_year_corrected`
   - `birth_month_corrected`
   - `birth_day_corrected`
   - `entry_year_corrected`
   - `entry_month_corrected`
   - `entry_day_corrected`
5. Majority correction may rewrite birth corrected fields.
6. Institution validator reads corrected fields.
7. Export writes corrected fields only.

### Bugs Fixed

- Removes split/single corrected-field inconsistency.
- Prevents export from reading raw values.

## Corrected Status-Field Lifecycle

### Purpose

Status fields should explain DateEngine outcomes.

### Date Status Fields

```python
birth_date_status
entry_date_status
```

Should include:

- missing component
- invalid component
- date does not exist
- unrecognized format
- future birth date
- entry cutoff violation
- age warning
- entry-before-birth warning

### Validation Fields

```python
_validation_ok
_validation_status
```

Should include institution/report validation outcomes.

### What Should Not Happen

- Institution validation should not overwrite DateEngine date status.
- Export should not write or modify status.
- Internal logic should not rely only on display text.

### Bugs Fixed

- Avoids brittle status-text logic.
- Keeps parser/business status distinct from report validation.

## Corrected Export Contract

### Purpose

Export should write finalized corrected fields only.

### Inputs

Normalized rows after DateEngine, majority correction, and validation.

### Outputs

Excel export workbook.

### Contract

```python
"ShnatLida" -> "birth_year_corrected"
"HodeshLida" -> "birth_month_corrected"
"YomLida" -> "birth_day_corrected"
"shnatknisa" -> "entry_year_corrected"
"Hodeshknisa" -> "entry_month_corrected"
"YomKnisa" -> "entry_day_corrected"
```

### What Should Not Happen

- Export should not parse dates.
- Export should not validate dates.
- Export should not fall back to original date fields when corrected fields are intentionally blank.

### Bugs Fixed

- Prevents data loss and invalid value resurrection during export.
- Aligns Web export and Dataset export behavior.

## Corrected Majority-Correction Flow

### Purpose

Correct likely wrong centuries for auto-completed birth years using sheet-level evidence.

### Inputs

Rows after row-level date normalization:

```python
[
    {"birth_year_corrected": 1934, "_birth_year_auto_completed": True},
    {"birth_year_corrected": 1935, "_birth_year_auto_completed": True},
    {"birth_year_corrected": 2024, "_birth_year_auto_completed": True}
]
```

### Ownership

- `date_standardization.apply_birth_year_majority_correction`
- Called by `standardization_pipeline._apply_birth_year_majority_correction`

### Correct Behavior

1. Include every row whose birth year was auto-completed.
2. Exclude explicit four-digit years.
3. Count 1900s and 2000s among eligible rows.
4. Apply approved correction rule.
5. Revalidate changed rows.
6. Rewrite corrected fields and status.
7. Preserve or log enough metadata for debugging.
8. Remove internal metadata only at final presentation/export boundary if needed.

### Output Example

```python
[
    {"birth_year_corrected": 1934, "birth_date_status": ""},
    {"birth_year_corrected": 1935, "birth_date_status": ""},
    {"birth_year_corrected": 1924, "birth_date_status": ""}
]
```

### Bugs Fixed

- Single-date numeric and separated inputs are eligible.
- Revalidation happens after century correction.
- Majority behavior becomes explainable and testable.

## Recommended Architecture

Recommended architecture additions:

1. `DateInput`
   - describes source kind and raw values.
2. Extended `DateParseResult`
   - stores structured validity, status, severity, and year metadata.
3. Shared corrected-field builder
   - applies one correction policy for all parser paths.
4. Explicit processing date/reference year
   - keeps the runtime year business-significant while making it explicit, consistent, logged, and testable.
5. Shared business rule definitions
   - keeps DateEngine and institution validator aligned.

Recommended dependency direction:

```text
web service
  -> extractor
  -> pipeline
  -> date adapter
  -> DateEngine
  -> date adapter writes corrected fields
  -> majority correction
  -> institution validator
  -> export writer
```

DateEngine should not depend on web, export, or validation modules.

## Implementation Order

1. Add tests that capture desired corrected behavior.
2. Extend `DateParseResult` compatibly.
3. Add explicit reference date/year handling.
4. Fix split source routing and partial split handling.
5. Make all two-digit expansion paths write metadata.
6. Create one corrected-field builder and use it for split and single flows.
7. Update majority correction eligibility and revalidation.
8. Align business validation semantics.
9. Align institution validation with corrected fields.
10. Align export behavior so corrected fields are the only date source.
11. Remove or deprecate unsafe fallback behavior.

## Exact Files / Functions To Change

### `src/excel_standardization/data_types.py`

- `DateParseResult`
  - Add structured fields while preserving existing public attributes.

### `src/excel_standardization/engines/date_engine.py`

- `parse_date`
  - Accept/derive explicit source kind.
  - Avoid routing partial split dates to single parsing.
- `parse_from_split_columns`
  - Support partial components.
  - Preserve valid components.
  - Report missing/invalid components.
- `parse_date_value`
  - Keep as single-value dispatcher.
  - Avoid unsafe integer serial assumptions where possible.
- `_parse_numeric_date_string`
  - Mark two-digit year auto-completion.
- `_parse_separated_date_string`
  - Use explicit reference year.
  - Mark auto-completion/defaulted year.
- `_parse_mixed_month_numeric`
  - Mark two-digit year auto-completion.
- `_expand_two_digit_year`
  - Use explicit reference year.
- `_validate_date`
  - Feed structured validity fields.
- `validate_business_rules`
  - Use explicit reference date.
  - Separate business validity and warnings.
- `validate_entry_before_birth`
  - Use structured validity semantics.

### `src/excel_standardization/processing/date_standardization.py`

- `detect_date_format_pattern`
  - Keep sheet-level pattern detection, possibly add confidence metadata later.
- `normalize_date_field`
  - Route any split field to split mode.
  - Build explicit DateEngine input.
  - Use shared corrected-field builder.
- `date_corrected_components`
  - Replace or expand into a complete corrected-field policy.
- `apply_birth_year_majority_correction`
  - Include all auto-completed birth years.
  - Revalidate changed rows.
  - Preserve/debug metadata intentionally.

### `src/excel_standardization/processing/standardization_pipeline.py`

- `normalize_dataset`
  - Provide reference date/year.
  - Keep ordering: detect pattern, normalize rows, majority correction, institution validation.
- `_apply_birth_year_majority_correction`
  - Continue delegating, but ensure corrected implementation is called.

### `src/excel_standardization/validation/institution_report_validator.py`

- `_validate_birth_date`
  - Align thresholds with DateEngine.
  - Read corrected fields first.
- `_validate_entry_date`
  - Align entry cutoff semantics.
  - Avoid re-parsing raw date strings.

### `webapp/services/export_schema.py`

- `EXPORT_MAPPING`
  - Keep mapping date headers to corrected fields only.

### `webapp/services/export_writer.py`

- `write_export_workbook`
  - Keep as pure writer.
  - Do not add date parsing or fallback.

### `src/excel_standardization/export/export_engine.py`

- `_map_row_to_export_fields`
  - Align with Web export.
  - Remove or constrain fallback from corrected date fields to original date fields for active flow.

## Safe Refactors

1. Add `DateInput` as an internal helper object.
2. Extend `DateParseResult` without removing existing fields.
3. Add status codes while keeping existing status text.
4. Introduce `reference_date` with a default at pipeline construction.
5. Add a single corrected-field builder.
6. Add tests before changing parser behavior.
7. Keep current export field names unchanged.
8. Keep original raw fields unchanged.

## Dangerous Refactors

1. Rewriting the entire DateEngine parser in one step.
2. Removing Excel serial support without source-format metadata.
3. Moving validation into export.
4. Changing majority correction directionality without business approval.
5. Changing Hebrew status text without updating UI/tests.
6. Removing original fields from rows.
7. Making institution validator the only source of date validation.

## Migration Risks

1. Existing tests may assert current imperfect behavior.
2. UI may depend on exact status text.
3. Export consumers may expect current fallback behavior.
4. Some files may contain integer years that are currently treated as Excel serials.
5. Majority correction may affect more rows after metadata is fixed.
6. Reference-year behavior must be agreed before changing production output.
7. Blank-vs-partial corrected component policy affects validation and export.

## Suggested Tests

### Split Date Tests

- Complete split birth date with four-digit year.
- Complete split birth date with two-digit year.
- Split entry date with two-digit year.
- Split date with invalid month.
- Split date with invalid day.

### Partial Split Date Tests

- Missing day preserves year/month.
- Missing month preserves year/day.
- Missing year preserves month/day.
- Non-numeric component blanks only that component.
- Partial split does not become empty single date.

### Single Date Tests

- Single `birth_date` separated DD/MM.
- Single `birth_date` separated MM/DD.
- Single `entry_date` separated DD/MM.
- Empty birth date behavior.
- Empty entry date behavior.

### Numeric String Tests

- `01022024`.
- `010224`.
- `1225`.
- `2020` as year-only partial date.
- Invalid numeric length.

### Separated Date Tests

- `01/02/24` with DD/MM.
- `01/02/24` with MM/DD.
- Dot-separated date.
- Two-part date with explicit reference year.
- Non-numeric separated date.

### Excel Serial / Type Tests

- Confirmed Excel serial date.
- Plain integer year should not become serial without metadata.
- Python `date`.
- Python `datetime`.

### Two-Digit Year Tests

- Reference year 2026:
  - `24 -> 2024`
  - `26 -> 2026`
  - `27 -> 1927`
- Every parser path marks auto-completion metadata.

### Business Validation Tests

- Future birth date.
- Entry date in current/reference year.
- Entry date after cutoff.
- Birth year before 1906.
- Age over 100 warning/error according to chosen policy.

### Majority Correction Tests

- Split two-digit years included.
- Single separated two-digit years included.
- Numeric compact two-digit years included.
- Explicit four-digit years excluded.
- Revalidation after correction.

### Institution Validation Tests

- Reads corrected fields first.
- Does not re-parse raw date strings.
- Writes `_validation_ok`.
- Writes `_validation_status`.
- Does not overwrite date status fields.

### Export Tests

- Web export writes corrected birth date components.
- Web export writes corrected entry date components.
- Export does not fall back to original date fields when corrected field is blank.
- Dataset `ExportEngine` matches Web export date behavior.
