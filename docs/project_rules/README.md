# Project Rules

This folder documents the project as it exists now: web API, upload flow,
workbook/session flow, extraction, pipeline, validation, export, UI/grid,
manual edits, runtime file safety, packaging, testing, security, and legacy
paths.

These documents are descriptive first. They separate:

- Approved behavior
- Current behavior
- Known limitations
- Needs approval

They are intended to help maintainers and AI agents keep changes aligned with
the actual codebase.

## Documents

- `PROJECT_ARCHITECTURE_RULES.md`
- `API_RULES.md`
- `UPLOAD_RULES.md`
- `WORKBOOK_SESSION_RULES.md`
- `WORKBOOK_LOADER_RULES.md`
- `EXCEL_EXTRACTION_RULES.md`
- `HEADER_DETECTION_RULES.md`
- `STANDARDIZATION_PIPELINE_RULES.md`
- `VALIDATION_RULES.md`
- `EXPORT_SYSTEM_RULES.md`
- `FRONTEND_GRID_RULES.md`
- `MANUAL_EDITS_RULES.md`
- `RUNTIME_FILE_SAFETY_RULES.md`
- `PACKAGING_RULES.md`
- `TESTING_RULES.md`
- `SECURITY_DEPENDENCY_RULES.md`
- `LEGACY_DISABLED_PATHS_RULES.md`

## Source Areas

- `webapp/`
- `src/excel_standardization/`
- `tests/`
- `installer/`
- root packaging files such as `pyproject.toml` and `ExcelNormalization.spec`

