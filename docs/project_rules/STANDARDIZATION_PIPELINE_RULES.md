# Standardization Pipeline Rules

## Purpose

Document how the engine pipeline normalizes rows and coordinates the
standardization flow.

## Scope

- `src/excel_standardization/processing/standardization_pipeline.py`
- engine modules under `src/excel_standardization/engines/`
- `src/excel_standardization/processing/*_standardization.py`
- `webapp/services/standardization_service.py`

## Main Files

- `src/excel_standardization/processing/standardization_pipeline.py`
- `src/excel_standardization/processing/name_standardization.py`
- `src/excel_standardization/processing/gender_standardization.py`
- `src/excel_standardization/processing/date_standardization.py`
- `src/excel_standardization/processing/identifier_standardization.py`
- `webapp/services/standardization_service.py`

## Responsibilities

- Keep original values immutable.
- Write standardized values into corrected fields.
- Apply engine-specific rules in a consistent order.
- Collect statistics and failure metadata.

## Data Flow

1. A `SheetDataset` enters the pipeline.
2. Name, gender, date, and identifier rules run.
3. Dataset-level patterns are detected where needed.
4. A normalized `SheetDataset` is returned.
5. Workbook-level validation runs after normalization in the web flow.

## Contracts

- Original values must never be overwritten in place.
- Corrected fields and status fields are separate outputs.
- Pipeline output must remain suitable for UI/grid and export consumers.

## What Must Never Change

- Approved engine behavior must not change as a side effect of orchestration
  refactors.
- Partial failures should remain isolated where current behavior expects it.

## Current Behavior

- The pipeline copies the row data before mutating corrected fields.
- Dataset-level name pattern detection is computed once per dataset.
- Manual edits are replayed after standardization in the web service.
- The normalized-row contract helper centralizes source/corrected/status field
  selection used by validation and export.

## Known Limitations

- Some orchestration details are still embedded in service code.
- The corrected-row contract is convention-based rather than a separate schema
  object.

## Tests That Should Cover It

- pipeline unit tests
- engine-specific regression tests
- service-level normalization tests
- dataset pattern and corrected-field visibility tests

## Open Questions / Future Improvements

- Whether to extract a shared normalized-row contract helper.
- Whether to reduce coupling between service orchestration and pipeline
  internals.
