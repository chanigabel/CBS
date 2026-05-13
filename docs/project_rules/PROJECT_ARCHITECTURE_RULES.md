# Project Architecture Rules

## Purpose

Document the top-level architecture of the project so future changes do not
break the contract between upload, extraction, standardization, validation,
UI/grid, manual edits, and export.

## Scope

- FastAPI web layer
- session-backed workbook state
- extraction and normalization pipeline
- institution validation
- export system
- static browser UI
- packaging/runtime folders
- legacy disabled paths

## Main Files

- `webapp/app.py`
- `webapp/api/*.py`
- `webapp/services/*.py`
- `src/excel_standardization/io_layer/*.py`
- `src/excel_standardization/processing/*.py`
- `src/excel_standardization/validation/institution_report_validator.py`
- `src/excel_standardization/export/export_engine.py`
- `webapp/static/js/*.js`
- `installer/Excelstandardization.iss`

## Responsibilities

- The web layer owns HTTP, session lookup, and UI payload shaping.
- The `src/excel_standardization` layer owns extraction, normalization, and
  validation logic.
- Export writes a new workbook from the in-memory normalized dataset.
- The browser UI is a consumer of API payloads, not the source of truth.

## Data Flow

1. Upload receives an Excel file.
2. The workbook is stored to `uploads/` and a working copy to `work/`.
3. The workbook is extracted into `WorkbookDataset` / `SheetDataset`.
4. The standardization pipeline adds corrected fields and statuses.
5. Workbook validation adds validation status fields.
6. The grid shows a shaped view of the in-memory rows.
7. Export writes a new `.xlsx` file to `output/`.

## Contracts

- Original input values remain immutable.
- Corrected values are written to `*_corrected` fields.
- Export uses the corrected fields for standardized columns.
- Row identity uses `_row_uid`.
- Hidden/internal helper fields stay internal unless a service explicitly
  exposes them.

## What Must Never Change

- Source workbooks are never modified in place.
- Export must remain a generated file, not a rewrite of the source file.
- Manual edits are session-backed and row-UID-based.
- Legacy disabled paths must stay disabled unless explicitly re-enabled.

## Current Behavior

- The active web path is the supported runtime path.
- The legacy direct-Excel orchestrator is disabled.
- `.xlsx`, `.xlsm`, and `.xls` are supported at the web/session layer.
- Static browser JS renders the grid and edit controls.

## Known Limitations

- The codebase does not have a separate frontend build system.
- Some responsibilities are still duplicated across services for pragmatic
  reasons.
- Excel compatibility still depends on the underlying reader/writer libraries.

## Tests That Should Cover It

- upload, workbook, standardization, validation, export, and edit integration
  tests
- end-to-end `.xlsx` and `.xls` flows
- regression tests for row identity, manual edits, and export output

## Open Questions / Future Improvements

- Whether to centralize workbook-loading dispatch into one helper.
- Whether to define a shared normalized-row schema contract module.
- Whether to shrink the shape-building logic in `WorkbookService`.

