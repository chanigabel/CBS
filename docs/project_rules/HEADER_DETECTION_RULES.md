# Header Detection Rules

## Purpose

Describe how source headers are located and mapped to field names.

## Scope

- `src/excel_standardization/io_layer/excel_reader.py`
- header/table helper modules under `src/excel_standardization/io_layer/`

## Main Files

- `src/excel_standardization/io_layer/excel_reader.py`
- `src/excel_standardization/io_layer/table_detection.py`
- `src/excel_standardization/io_layer/column_detection.py`
- `src/excel_standardization/io_layer/field_matching.py`

## Responsibilities

- Find the table region.
- Detect header rows and subheaders.
- Map Excel text to canonical source field names.

## Data Flow

1. The reader scans candidate rows.
2. Field keywords and merge information are normalized.
3. Column mappings are cached per worksheet.
4. The extractor consumes the resulting mapping.

## Contracts

- Header detection must ignore corrected/status output columns when a generated
  workbook is reloaded as source.
- The detected field order should match the Excel column order as closely as
  possible.

## What Must Never Change

- Re-uploading an exported workbook must not re-import output-only columns as
  source input.
- Internal helper columns must remain filtered from source detection.

## Current Behavior

- `ExcelReader` uses keyword-based heuristics, merged-cell awareness, and
  table-region scoring.
- The reader caches detection results per worksheet instance.

## Known Limitations

- Header detection is heuristic, not schema-driven.
- Complex spreadsheets may still need fixture-based regression coverage.

## Tests That Should Cover It

- debug fixtures for multi-row and merged headers
- extraction tests for unusual layouts
- regression tests for exported workbook re-imports

## Open Questions / Future Improvements

- Whether to introduce a more explicit header schema for known templates.
- Whether to separate "source header detection" from "output workbook
  sanitation" more strongly.
