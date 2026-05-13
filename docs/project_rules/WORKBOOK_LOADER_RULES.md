# Workbook Loader Rules

## Purpose

Describe the loader policy for `.xlsx`, `.xlsm`, and `.xls` files.

## Scope

- `webapp/services/upload_service.py`
- `webapp/services/workbook_service.py`
- `webapp/services/standardization_service.py`
- `webapp/services/export_service.py`
- `src/excel_standardization/io_layer/xls_reader.py`
- `src/excel_standardization/io_layer/excel_to_json_extractor.py`

## Main Files

- `src/excel_standardization/io_layer/xls_reader.py`
- `src/excel_standardization/io_layer/excel_to_json_extractor.py`
- `webapp/services/workbook_loader.py`
- `webapp/services/upload_service.py`
- `webapp/services/workbook_service.py`
- `webapp/services/standardization_service.py`
- `webapp/services/export_service.py`

## Responsibilities

- Route `.xls` to the legacy reader path.
- Route `.xlsx` / `.xlsm` to openpyxl-based extraction.
- Keep workbook loading consistent across upload, sheet view, standardize, and
  export.

## Data Flow

1. A session points to a working copy on disk.
2. The loader chooses a reader by file suffix.
3. The reader produces sheet names or a workbook dataset.
4. Services cache the dataset back into the session.

## Contracts

- Loader selection is suffix-based in the current code.
- `.xls` uses the `xls_reader` path.
- `.xlsx` and `.xlsm` use openpyxl-based loading.

## What Must Never Change

- Loader choice must not depend on UI state.
- Source files must not be overwritten during loading.

## Current Behavior

- Upload validates workbook openness before a session is created.
- Workbook and export services can lazily load the dataset if needed.
- The workbook service can lazy-load individual sheets when absent.
- `webapp/services/workbook_loader.py` is the canonical loader dispatch for
  supported workbook types.

## Known Limitations

- Loader selection is duplicated in multiple services.
- Content-based file type detection is not the current policy.

## Tests That Should Cover It

- lazy-load tests for `.xls`
- upload validation tests
- workbook summary and sheet-loading tests
- export lazy-load tests

## Open Questions / Future Improvements

- Whether to collapse the remaining helper aliases into a smaller loader API.
- Whether to treat uppercase suffixes or mismatched extensions more
  explicitly.
