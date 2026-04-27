# Implementation Plan: Excel Normalization Web App

## Overview

Build a local FastAPI web application on top of the existing Excel normalization pipeline. The implementation proceeds in layers: repository cleanup → project dependencies → data models → service layer → API layer → frontend → tests. Each task builds on the previous and ends with all components wired together.

## Tasks

- [x] 1. Repository cleanup and dependency setup
  - [x] 1.1 Move root-level ad-hoc scripts to `scripts/` folder
    - Move `run_normalization_b.py`, `run_parity_python.py`, `run_processors_b.py`, `compare_workbooks.py`, `compare_workbooks_structural.py`, `debug_date_groups_b.py`, `debug_structural_parity.py` from project root into `scripts/`
    - _Requirements: 10.4_

  - [x] 1.2 Update `pyproject.toml` to add FastAPI and Uvicorn dependencies
    - Add `fastapi>=0.100.0` and `uvicorn[standard]>=0.23.0` to `[project.dependencies]`
    - Add `python-multipart>=0.0.6` (required for FastAPI file uploads)
    - Add `jinja2>=3.1.0` (required for Jinja2 templates)
    - _Requirements: 10.6_

  - [x] 1.3 Add `uploads/`, `work/`, and `output/` to `.gitignore`
    - Ensure the three runtime directories are gitignored so no user data is committed
    - _Requirements: 9.2_

- [x] 2. Webapp package skeleton
  - [x] 2.1 Create `webapp/` package with all sub-package `__init__.py` files
    - Create `webapp/__init__.py`, `webapp/api/__init__.py`, `webapp/services/__init__.py`, `webapp/models/__init__.py`
    - Create empty placeholder files for all modules listed in the design so imports resolve: `webapp/api/upload.py`, `webapp/api/workbook.py`, `webapp/api/normalize.py`, `webapp/api/edit.py`, `webapp/api/export.py`
    - Create empty placeholder files: `webapp/services/session_service.py`, `webapp/services/upload_service.py`, `webapp/services/workbook_service.py`, `webapp/services/normalization_service.py`, `webapp/services/edit_service.py`, `webapp/services/export_service.py`
    - Create `webapp/models/session.py`, `webapp/models/requests.py`, `webapp/models/responses.py`
    - Create `webapp/templates/` and `webapp/static/` directories (add `.gitkeep` if empty)
    - _Requirements: 10.1, 10.2_

- [x] 3. Pydantic models and SessionRecord
  - [x] 3.1 Implement `SessionRecord` dataclass in `webapp/models/session.py`
    - Fields: `session_id: str`, `source_file_path: str`, `working_copy_path: str`, `original_filename: str`, `status: str`, `workbook_dataset: Optional[WorkbookDataset]`, `edits: dict`
    - _Requirements: 2.1, 2.2_

  - [x] 3.2 Implement Pydantic request models in `webapp/models/requests.py`
    - `CellEditRequest`: `row_index: int`, `field_name: str`, `new_value: str`
    - _Requirements: 6.1_

  - [x] 3.3 Implement Pydantic response models in `webapp/models/responses.py`
    - `UploadResponse`: `session_id: str`, `sheet_names: list[str]`
    - `SheetSummary`: `sheet_name: str`, `row_count: int`, `field_names: list[str]`
    - `WorkbookSummary`: `session_id: str`, `sheets: list[SheetSummary]`
    - `SheetDataResponse`: `sheet_name: str`, `field_names: list[str]`, `rows: list[dict]`
    - `PerSheetStat`: `sheet_name: str`, `rows: int`, `success_rate: float`
    - `NormalizeResponse`: `session_id: str`, `status: str`, `sheets_processed: int`, `total_rows: int`, `per_sheet_stats: list[PerSheetStat]`
    - `CellEditResponse`: `row_index: int`, `updated_row: dict`
    - `ErrorResponse`: `detail: str`
    - _Requirements: 1.3, 3.2, 4.2, 5.5, 6.3_

- [x] 4. Service layer — SessionService
  - [x] 4.1 Implement `SessionService` in `webapp/services/session_service.py`
    - Module-level `_registry: dict[str, SessionRecord] = {}` as the shared in-memory store
    - `create(record: SessionRecord) -> None`
    - `get(session_id: str) -> SessionRecord` — raises `HTTPException(404)` if not found
    - `update(session_id: str, **kwargs) -> None`
    - `delete(session_id: str) -> None`
    - _Requirements: 2.1, 2.3, 2.4, 2.5_

  - [x]* 4.2 Write unit tests for `SessionService` in `tests/webapp/test_session_service.py`
    - Test `create` then `get` returns the same record
    - Test `get` with unknown session_id raises HTTPException 404
    - Test `update` mutates the stored record fields
    - _Requirements: 2.3, 2.5_

- [x] 5. Service layer — UploadService
  - [x] 5.1 Implement `UploadService` in `webapp/services/upload_service.py`
    - Constructor: `__init__(self, session_service: SessionService, uploads_dir: Path, work_dir: Path)`
    - `handle_upload(filename: str, file_bytes: bytes) -> UploadResponse`
      - Validate extension is `.xlsx` or `.xlsm`; raise `HTTPException(400)` if not
      - Generate `session_id = str(uuid4())`
      - Write bytes to `uploads/{session_id}{ext}` (source file, never modified)
      - Copy to `work/{session_id}{ext}` (working copy)
      - Call `ExcelToJsonExtractor().extract_workbook_to_json(working_copy_path)` to validate and extract `WorkbookDataset`; raise `HTTPException(422)` if it fails
      - Create `SessionRecord` with `status="uploaded"` and store via `session_service.create`
      - Return `UploadResponse(session_id=..., sheet_names=[...])`
    - _Requirements: 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 2.2_

  - [x]* 5.2 Write property test for upload preserves source file integrity (`Property 1`)
    - **Property 1: Upload preserves source file integrity**
    - **Validates: Requirements 1.2, 9.2**
    - Use `@given(st.binary(min_size=1))` with a valid `.xlsx` wrapper strategy
    - After `handle_upload`, assert bytes in `uploads/` are identical to input bytes
    - Assert a separate copy exists in `work/`
    - Place in `tests/webapp/test_webapp_properties.py`

  - [x]* 5.3 Write property test for upload response reflects workbook structure (`Property 2`)
    - **Property 2: Upload response reflects actual workbook structure**
    - **Validates: Requirements 1.3, 1.6**
    - Use a `workbook_strategy()` Hypothesis strategy that generates valid `openpyxl.Workbook` instances with random sheet names
    - Assert `session_id` in response is a valid UUID
    - Assert `sheet_names` in response exactly matches the workbook's sheet names
    - Place in `tests/webapp/test_webapp_properties.py`

  - [x]* 5.4 Write property test for invalid extensions always rejected (`Property 3`)
    - **Property 3: Invalid file extensions are always rejected**
    - **Validates: Requirements 1.4**
    - Use `@given(st.text().filter(lambda x: not x.lower().endswith(('.xlsx', '.xlsm'))))`
    - Assert `handle_upload` raises `HTTPException` with status 400 for any non-xlsx/xlsm extension
    - Place in `tests/webapp/test_webapp_properties.py`

  - [x]* 5.5 Write unit tests for `UploadService` in `tests/webapp/test_upload_service.py`
    - Test valid `.xlsx` upload creates session and returns correct sheet names
    - Test `.xlsm` extension is accepted
    - Test `.csv` extension raises HTTPException 400
    - Test corrupted bytes raises HTTPException 422
    - _Requirements: 1.1, 1.4, 1.5_

- [x] 6. Service layer — WorkbookService
  - [x] 6.1 Implement `WorkbookService` in `webapp/services/workbook_service.py`
    - Constructor: `__init__(self, session_service: SessionService)`
    - `get_summary(session_id: str) -> WorkbookSummary`
      - Get session; read `workbook_dataset`; build `WorkbookSummary` with per-sheet row counts and field names
    - `get_sheet_data(session_id: str, sheet_name: str) -> SheetDataResponse`
      - Get session; find sheet by name in `workbook_dataset`; raise `HTTPException(404)` if not found
      - Return `SheetDataResponse` with `field_names` and `rows`
    - _Requirements: 3.1, 3.2, 3.3, 4.1, 4.2, 4.3, 4.4, 4.5_

  - [x]* 6.2 Write unit tests for `WorkbookService` in `tests/webapp/test_workbook_service.py`
    - Test `get_summary` returns correct sheet names, row counts, and field names
    - Test `get_sheet_data` returns rows for a valid sheet
    - Test `get_sheet_data` raises HTTPException 404 for unknown sheet name
    - _Requirements: 3.2, 4.4_

- [x] 7. Service layer — NormalizationService
  - [x] 7.1 Implement `NormalizationService` in `webapp/services/normalization_service.py`
    - Constructor: `__init__(self, session_service: SessionService)`
    - `normalize(session_id: str) -> NormalizeResponse`
      - Get session
      - Instantiate `NormalizationOrchestrator()` fresh per request
      - Call `orchestrator.process_workbook_json(working_copy_path, temp_output_path)`
      - Re-extract normalized data: call `ExcelToJsonExtractor().extract_workbook_to_json(working_copy_path)` then `NormalizationPipeline().normalize_dataset(sheet)` for each sheet
      - Collect per-sheet stats from `sheet.metadata` (success rates)
      - Update `session.workbook_dataset` with normalized sheets; set `session.status = "normalized"`
      - If all sheets fail, raise `HTTPException(500)`; otherwise return `NormalizeResponse` with stats
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7, 5.8, 12.1, 12.2_

  - [x]* 7.2 Write property test for normalization updates session status and dataset (`Property 6`)
    - **Property 6: Normalization updates session status and dataset**
    - **Validates: Requirements 5.4**
    - Use `workbook_strategy()` to generate valid workbooks; upload and then normalize
    - Assert `session.status == "normalized"` after normalize call
    - Assert `session.workbook_dataset` rows contain `_corrected` fields
    - Place in `tests/webapp/test_webapp_properties.py`

  - [x]* 7.3 Write property test for source file never modified (`Property 7`)
    - **Property 7: Source file is never modified**
    - **Validates: Requirements 5.8, 7.5, 9.2**
    - Use `workbook_strategy()`; record SHA-256 of `uploads/` file after upload
    - Run normalize, then export; re-read `uploads/` file and assert hash unchanged
    - Place in `tests/webapp/test_webapp_properties.py`

  - [x]* 7.4 Write property test for web app normalization equivalence (`Property 9`)
    - **Property 9: Web app normalization equivalence**
    - **Validates: Requirements 12.5**
    - Use `@settings(max_examples=50)` due to full pipeline cost
    - For a given workbook, run `NormalizationService.normalize` and also run `NormalizationOrchestrator.process_workbook_json` + `ExcelToJsonExtractor` directly
    - Assert the resulting normalized rows are equivalent
    - Place in `tests/webapp/test_webapp_properties.py`

  - [x]* 7.5 Write unit tests for `NormalizationService` in `tests/webapp/test_normalization_service.py`
    - Test successful normalization sets status to "normalized" and returns stats
    - Test partial sheet failure logs error and continues (returns 200 with stats)
    - Test all-sheets failure raises HTTPException 500
    - _Requirements: 5.4, 5.6, 5.7_

- [x] 8. Service layer — EditService and ExportService
  - [x] 8.1 Implement `EditService` in `webapp/services/edit_service.py`
    - Constructor: `__init__(self, session_service: SessionService)`
    - `edit_cell(session_id: str, sheet_name: str, req: CellEditRequest) -> CellEditResponse`
      - Get session; find sheet; raise `HTTPException(404)` if sheet not found
      - Validate `row_index` in `[0, len(rows)-1]`; raise `HTTPException(400)` if out of range
      - Validate `field_name` exists in row keys; raise `HTTPException(400)` if not found
      - Mutate `rows[row_index][field_name] = new_value`
      - Record edit in `session.edits[(sheet_name, row_index, field_name)] = new_value`
      - Return `CellEditResponse(row_index=..., updated_row=rows[row_index])`
    - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5, 6.6_

  - [x] 8.2 Implement `ExportService` in `webapp/services/export_service.py`
    - Constructor: `__init__(self, session_service: SessionService, output_dir: Path)`
    - `export(session_id: str) -> Path`
      - Get session
      - Build output filename: `{original_stem}_normalized_{timestamp}.xlsx`
      - Call `ExportEngine().export_from_normalized_dataset(session.workbook_dataset, output_path)`
      - On failure, catch exception, log it, and raise `HTTPException(500)` — do NOT modify session state
      - Return `output_path`
    - _Requirements: 7.1, 7.2, 7.4, 7.5, 7.6, 11.4_

  - [x]* 8.3 Write property test for cell edit round-trip (`Property 8`)
    - **Property 8: Cell edit round-trip**
    - **Validates: Requirements 6.2, 6.3**
    - Use `sheet_dataset_strategy()` to generate `SheetDataset` instances with random field names and rows
    - After `edit_cell`, call `get_sheet_data` and assert `rows[row_index][field_name] == new_value`
    - Place in `tests/webapp/test_webapp_properties.py`

  - [x]* 8.4 Write unit tests for `EditService` in `tests/webapp/test_edit_service.py`
    - Test valid edit mutates the in-memory row and returns updated row
    - Test out-of-range `row_index` raises HTTPException 400
    - Test unknown `field_name` raises HTTPException 400
    - Test edit is recorded in `session.edits`
    - _Requirements: 6.2, 6.3, 6.4, 6.5, 6.6_

  - [x]* 8.5 Write unit tests for `ExportService` in `tests/webapp/test_export_service.py`
    - Test successful export returns a valid file path with `_normalized` suffix
    - Test export failure raises HTTPException 500 and does not modify session state
    - _Requirements: 7.2, 7.4, 7.6, 11.4_

- [x] 9. Checkpoint — Ensure all service layer tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 10. API layer — all routers
  - [x] 10.1 Implement `POST /api/upload` router in `webapp/api/upload.py`
    - Accept `file: UploadFile = File(...)` via multipart form
    - Read file bytes; call `UploadService.handle_upload(file.filename, await file.read())`
    - Return `UploadResponse`; let HTTPExceptions propagate as-is
    - _Requirements: 1.1, 1.4, 1.5_

  - [x] 10.2 Implement workbook routers in `webapp/api/workbook.py`
    - `GET /api/workbook/{session_id}/summary` → calls `WorkbookService.get_summary`
    - `GET /api/workbook/{session_id}/sheet/{sheet_name}` → calls `WorkbookService.get_sheet_data`
    - _Requirements: 3.1, 4.1_

  - [x] 10.3 Implement `POST /api/workbook/{session_id}/normalize` router in `webapp/api/normalize.py`
    - Calls `NormalizationService.normalize(session_id)`; returns `NormalizeResponse`
    - _Requirements: 5.1_

  - [x] 10.4 Implement `PATCH /api/workbook/{session_id}/sheet/{sheet_name}/cell` router in `webapp/api/edit.py`
    - Accepts `CellEditRequest` body; calls `EditService.edit_cell`; returns `CellEditResponse`
    - _Requirements: 6.1_

  - [x] 10.5 Implement `POST /api/workbook/{session_id}/export` router in `webapp/api/export.py`
    - Calls `ExportService.export(session_id)`; returns `FileResponse` with `Content-Disposition: attachment`
    - _Requirements: 7.1, 7.3_

  - [x]* 10.6 Write unit tests for API upload endpoint in `tests/webapp/test_api_upload.py`
    - Use FastAPI `TestClient`; test valid upload returns 200 with session_id and sheet_names
    - Test invalid extension returns 400; test corrupted file returns 422
    - _Requirements: 1.1, 1.4, 1.5_

  - [x]* 10.7 Write unit tests for API workbook endpoints in `tests/webapp/test_api_workbook.py`
    - Test summary endpoint returns correct structure for valid session
    - Test sheet data endpoint returns rows for valid sheet
    - Test 404 for unknown session and unknown sheet
    - _Requirements: 3.1, 4.1, 4.4_

  - [x]* 10.8 Write unit tests for normalize, edit, and export API endpoints in `tests/webapp/test_api_normalize.py`, `test_api_edit.py`, `test_api_export.py`
    - Test normalize returns 200 with stats for valid session
    - Test edit returns 200 with updated row; test 400 for bad row_index/field_name
    - Test export returns file download response
    - _Requirements: 5.1, 6.1, 7.1_

- [x] 11. FastAPI app entry point
  - [x] 11.1 Implement `webapp/app.py` — the FastAPI application entry point
    - Create `FastAPI` app instance with title
    - Mount `webapp/static/` at `/static`
    - Register `Jinja2Templates` from `webapp/templates/`
    - Include all five routers with `prefix="/api"`
    - Add `GET /` route that renders `index.html`
    - On startup: create `uploads/`, `work/`, `output/` directories if they don't exist
    - Configure basic logging
    - _Requirements: 8.1, 9.1, 10.1_

  - [x]* 10.9 Write property test for non-existent session always returns 404 (`Property 4`)
    - **Property 4: Non-existent session always returns 404**
    - **Validates: Requirements 2.3**
    - Use `@given(st.uuids())` with `TestClient`
    - Assert all session-scoped endpoints return 404 for any UUID not in the registry
    - Place in `tests/webapp/test_webapp_properties.py`

  - [x]* 10.10 Write property test for session initialization invariant (`Property 5`)
    - **Property 5: Session initialization invariant**
    - **Validates: Requirements 2.2**
    - Use `workbook_strategy()`; after upload assert `session.status == "uploaded"`, paths are non-empty, `workbook_dataset` has at least one sheet
    - Place in `tests/webapp/test_webapp_properties.py`

- [x] 12. Frontend — HTML template and CSS
  - [x] 12.1 Implement `webapp/templates/index.html` — the single-page UI skeleton
    - HTML5 document with `<head>` linking to `/static/style.css` and `/static/app.js`
    - Sections: `#upload-section` (file input + submit button), `#sheet-selector` (hidden initially), `#grid-section` (hidden initially), `#action-bar` (Normalize + Export buttons), `#error-banner` (hidden initially)
    - No external CDN links — all assets served locally
    - _Requirements: 8.1, 8.2, 8.3, 9.5_

  - [x] 12.2 Implement `webapp/static/style.css` — layout and grid styles
    - Base layout: centered container, responsive width
    - Grid table styles: scrollable wrapper, alternating row colors
    - Corrected cell highlight: light green background for `_corrected` fields that differ from original
    - Error banner: red/orange dismissible bar at top of page
    - Inline edit input styles
    - _Requirements: 8.5, 8.7, 8.10_

- [x] 13. Frontend — JavaScript application logic
  - [x] 13.1 Implement `webapp/static/app.js` — upload flow and sheet selector
    - `handleUpload(event)`: reads file, calls `POST /api/upload`, stores `sessionId` in module state, renders sheet selector tabs/dropdown from `sheet_names`
    - `loadSheet(sheetName)`: calls `GET /api/workbook/{sessionId}/sheet/{sheetName}`, stores current sheet data, calls `renderGrid()`
    - Error handling: on any non-2xx response, extract `detail` field and call `showError(message)`
    - `showError(message)` / `dismissError()`: show/hide `#error-banner`
    - _Requirements: 8.3, 8.4, 8.5, 8.10_

  - [x] 13.2 Implement grid rendering and inline cell editing in `webapp/static/app.js`
    - `renderGrid(sheetData)`: builds `<table>` from `field_names` and `rows`; for each row, render original columns and `_corrected` columns side-by-side with color coding (green if corrected value differs from original)
    - `makeEditable(cell, rowIndex, fieldName)`: on cell click, replace content with `<input>`; on blur/Enter call `PATCH .../cell` with `{row_index, field_name, new_value}`; on success update in-memory row and re-render cell
    - _Requirements: 8.5, 8.7, 8.8_

  - [x] 13.3 Implement normalize and export actions in `webapp/static/app.js`
    - `runNormalization()`: calls `POST /api/workbook/{sessionId}/normalize`; on success calls `loadSheet(currentSheet)` to refresh grid with corrected values
    - `exportWorkbook()`: calls `POST /api/workbook/{sessionId}/export`; triggers browser file download via a temporary `<a download>` element or `window.location`
    - Wire "Run Normalization" and "Export / Download" buttons to these functions
    - _Requirements: 8.6, 8.9_

- [x] 14. Checkpoint — Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 15. Integration test and README update
  - [x] 15.1 Write integration test for full upload → normalize → export workflow in `tests/webapp/test_integration_webapp.py`
    - Use FastAPI `TestClient` with a real sample `.xlsx` file from the repo
    - Assert upload returns valid session_id and sheet names
    - Assert normalize returns 200 with stats
    - Assert export returns a binary file response
    - Assert source file in `uploads/` is unchanged after the full workflow
    - _Requirements: 5.2, 5.3, 7.2, 9.2, 12.1_

  - [x] 15.2 Update `README.md` with webapp startup instructions
    - Add a "Web Application" section explaining how to install dependencies (`pip install -e ".[dev]"`) and start the server (`uvicorn webapp.app:app --reload`)
    - Document the default URL (`http://localhost:8000`)
    - _Requirements: 10.7_

- [x] 16. Final checkpoint — Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

## Notes

- Tasks marked with `*` are optional and can be skipped for a faster MVP
- The existing `src/excel_normalization/` code is never moved or modified
- `uploads/`, `work/`, and `output/` are runtime directories created at startup and gitignored
- Property tests use Hypothesis strategies: `workbook_strategy()` (generates valid openpyxl workbooks) and `sheet_dataset_strategy()` (generates SheetDataset instances)
- The server is started manually with: `uvicorn webapp.app:app --reload`
