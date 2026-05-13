# Agent: Header Detection

## Mission

Review and maintain source header detection and table-region detection logic.

## Files To Inspect First

- `docs/project_rules/HEADER_DETECTION_RULES.md`
- `src/excel_standardization/io_layer/excel_reader.py`
- `src/excel_standardization/io_layer/table_detection.py`
- `src/excel_standardization/io_layer/column_detection.py`
- `src/excel_standardization/io_layer/field_matching.py`

## Rules To Follow

- Keep detection heuristic changes narrow.
- Treat generated/exported workbook columns as non-source columns.

## What The Agent May Change

- Detection heuristics, cached lookups, regression fixtures, and docs when
  requested.

## What The Agent Must Not Change

- Business rules in the standardization engines.
- Source data mutation behavior.

## Required Tests

- header detection regression tests
- exported workbook re-import tests
- merged and multi-row header tests

## Regression Checklist

- source columns map correctly
- corrected/status columns are ignored on re-import
- caches do not stale out across worksheets

## Expected Output Format

- detection findings
- mapping findings
- cache findings
- missing tests

## Safety Constraints

- Keep the reader deterministic.
- Do not silently invent missing headers.

