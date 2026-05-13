# Export System Rules

## Purpose

Document how normalized rows are turned into downloadable Excel workbooks.

## Scope

- `webapp/services/export_service.py`
- `webapp/services/export_writer.py`
- `webapp/services/export_rows.py`
- `webapp/services/export_schema.py`
- `src/excel_standardization/export/export_engine.py`
- `src/excel_standardization/export/excel_safe.py`

## Main Files

- `webapp/services/export_service.py`
- `webapp/services/export_writer.py`
- `webapp/services/export_rows.py`
- `webapp/services/export_schema.py`
- `src/excel_standardization/export/export_engine.py`
- `src/excel_standardization/export/excel_safe.py`

## Responsibilities

- Build a valid `.xlsx` workbook.
- Export corrected standardized values.
- Preserve row and sheet ordering where the current contract requires it.
- Keep output safe for Excel to open.

## Data Flow

1. Session workbook dataset is loaded or reused.
2. Export rows are assembled from the normalized dataset.
3. Sheet titles and cell values are sanitized.
4. Workbook is written to `output/`.

## Contracts

- Export must use corrected standardized fields for standardized columns.
- Export must not depend on UI visibility.
- Output workbook must remain openable by Excel.

## What Must Never Change

- Source workbook values must not be modified.
- Corrected-only export mapping must remain deliberate and test-covered.
- Sheet/cell safety helpers must continue to protect workbook integrity.

## Current Behavior

- Active web export uses `EXPORT_MAPPING`.
- Compatibility `ExportEngine` maps its export headers separately but uses the
  same safety helpers.
- `safe_sheet_title` and `safe_cell_value` sanitize workbook output.

## Known Limitations

- Some schema and mapping logic is duplicated between export paths.
- Unsupported or unexpected value types are coerced into a safe written form.

## Tests That Should Cover It

- active export tests
- compatibility export tests
- workbook openability tests
- sheet-name collision tests

## Open Questions / Future Improvements

- Whether to factor a shared export assembly helper for both export paths.
- Whether to log coercions for unsupported cell types more explicitly.

