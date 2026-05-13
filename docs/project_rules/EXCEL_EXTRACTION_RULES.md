# Excel Extraction Rules

## Purpose

Document how Excel worksheets are converted into `WorkbookDataset` and
`SheetDataset` objects.

## Scope

- `src/excel_standardization/io_layer/excel_reader.py`
- `src/excel_standardization/io_layer/excel_to_json_extractor.py`
- `src/excel_standardization/io_layer/xls_reader.py`
- `src/excel_standardization/io_layer/table_detection.py`
- `src/excel_standardization/io_layer/column_detection.py`

## Main Files

- `src/excel_standardization/io_layer/excel_reader.py`
- `src/excel_standardization/io_layer/excel_to_json_extractor.py`
- `src/excel_standardization/io_layer/xls_reader.py`
- related IO helper modules in `src/excel_standardization/io_layer/`

## Responsibilities

- Detect the relevant table area.
- Detect source columns and header rows.
- Extract rows into JSON-like dictionaries.
- Preserve original cell values as read from Excel.

## Data Flow

1. Workbook is opened by the loader.
2. `ExcelReader` detects table region and column mapping.
3. `ExcelToJsonExtractor` converts rows into `JsonRow` objects.
4. The result becomes a `SheetDataset` or `WorkbookDataset`.

## Contracts

- Extraction must preserve original values.
- Missing or unreadable data should become `None` or error metadata, not a
  silent mutation.
- Formula handling must follow the extractor's current `data_only` and formula
  policy.

## What Must Never Change

- Extraction must not run business standardization rules.
- Source workbooks must remain unmodified.
- Empty or invalid header regions must not be guessed into fake source fields.

## Current Behavior

- `ExcelReader` handles complex header/table detection.
- `ExcelToJsonExtractor` can detect missing headers and return an empty dataset
  with error metadata.
- `.xls` extraction uses the dedicated reader path.

## Known Limitations

- Header detection is heuristic and may require regression tests for unusual
  workbook layouts.
- `.xls` and `.xlsx` use different underlying readers.

## Tests That Should Cover It

- header detection tests
- extraction tests for merged cells, multi-row headers, formulas, and empty
  sheets
- `.xls` regression fixtures

## Open Questions / Future Improvements

- Whether to expose a shared extraction contract module for all loaders.
- Whether to reduce duplicated header-detection heuristics across IO helpers.
