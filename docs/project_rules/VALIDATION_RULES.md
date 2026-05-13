# Validation Rules

## Purpose

Document how institution/Mosad validation works on normalized rows.

## Scope

- `src/excel_standardization/validation/institution_report_validator.py`
- `webapp/services/standardization_service.py`
- validation-related tests and API routes

## Main Files

- `src/excel_standardization/validation/institution_report_validator.py`
- `webapp/services/standardization_service.py`
- `tests/test_institution_report_validator.py`

## Responsibilities

- Validate institution report rows and workbook-wide constraints.
- Prefer corrected values where validation is meant to evaluate normalized
  output.
- Record validation status back into the rows.

## Data Flow

1. Normalized sheets are collected.
2. Workbook metadata is passed to the validator.
3. Row and workbook checks run.
4. `_validation_status` and `_validation_ok` are written back into rows.

## Contracts

- Validation should use the post-standardization values that downstream exports
  will use.
- Missing or invalid required values must be reported clearly.

## What Must Never Change

- Validation must not mutate source files.
- Validation must not erase corrected values.
- Workbook-level duplicate checks must remain workbook-wide.

## Current Behavior

- The validator checks MosadID, SugMosad, residence counts, names, identifier,
  gender, and date fields.
- It uses corrected fields for many checks and workbook metadata where
  available.

## Known Limitations

- Validation and message construction are tightly coupled in one module.
- Field selection policy is partly implicit.

## Tests That Should Cover It

- row-level validator tests
- workbook-wide duplicate tests
- mixed corrected/original field tests
- type-edge-case tests

## Open Questions / Future Improvements

- Whether to split pure validation from row mutation.
- Whether to centralize corrected-vs-original field selection policy.

