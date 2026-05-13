# Runtime File Safety Rules

## Purpose

Document how the project protects original files and writes safe runtime output.

## Scope

- upload/work/output folder usage
- export filename and workbook safety
- file handling in services
- installer-created runtime folders

## Main Files

- `webapp/services/upload_service.py`
- `webapp/services/export_service.py`
- `webapp/services/export_rows.py`
- `webapp/services/export_writer.py`
- `src/excel_standardization/export/excel_safe.py`
- `installer/Excelstandardization.iss`

## Responsibilities

- Preserve source files.
- Write working copies and exports to separate folders.
- Produce Excel-safe output values.
- Keep filenames valid for the runtime platform.

## Data Flow

1. Upload saves the original file to `uploads/`.
2. A working copy is saved to `work/`.
3. Export writes a new workbook to `output/`.
4. The installer/bootstrapper creates these folders if needed.

## Contracts

- The source file must never be modified in place.
- Exported workbooks must remain valid Excel files.
- Sanitization must not change business data beyond safe Excel coercion.

## What Must Never Change

- Original user files stay untouched.
- Exported files must be openable by Excel.
- Runtime paths must not cross directory boundaries unexpectedly.

## Current Behavior

- `safe_sheet_title` sanitizes worksheet names.
- `safe_cell_value` removes illegal characters and protects formula-like text.
- Export filenames are sanitized for filesystem safety.

## Known Limitations

- Coercion of unsupported object types may hide an upstream data issue unless
  the caller logs it.
- File-system and workbook safety still depend partly on platform behavior.

## Tests That Should Cover It

- filename sanitization tests
- workbook openability tests
- illegal character and sheet-title tests

## Open Questions / Future Improvements

- Whether to surface coercion warnings more visibly.
- Whether to standardize a shared path-safety helper for all file writes.

