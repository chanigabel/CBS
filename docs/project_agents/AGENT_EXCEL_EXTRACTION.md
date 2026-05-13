# Agent: Excel Extraction

## Mission

Review and maintain workbook-to-dataset extraction, including `.xls` and
`.xlsx/.xlsm` paths.

## Files To Inspect First

- `docs/project_rules/EXCEL_EXTRACTION_RULES.md`
- `src/excel_standardization/io_layer/excel_reader.py`
- `src/excel_standardization/io_layer/excel_to_json_extractor.py`
- `src/excel_standardization/io_layer/xls_reader.py`
- `src/excel_standardization/io_layer/table_detection.py`
- `src/excel_standardization/io_layer/column_detection.py`

## Rules To Follow

- Preserve original cell values.
- Keep header/table detection heuristic changes small and test-backed.

## What The Agent May Change

- Extraction helpers, detection helpers, tests, and docs when requested.

## What The Agent Must Not Change

- Standardization business logic.
- Export rules.

## Required Tests

- header detection tests
- multi-row and merged-header tests
- `.xls` regression tests
- extraction tests for formulas and empty cells

## Regression Checklist

- correct field mapping
- merged-cell handling
- no-header behavior
- `.xls` and `.xlsx` parity where supported

## Expected Output Format

- extraction path findings
- detection path findings
- tests run / missing

## Safety Constraints

- No runtime code changes unless explicitly requested.
- Do not guess field names when detection fails.

