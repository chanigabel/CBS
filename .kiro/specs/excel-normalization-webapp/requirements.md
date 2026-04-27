# Requirements Document

## Introduction

This feature adds a local web application layer on top of the existing Python Excel normalization project. The web app allows non-technical users to upload an Excel workbook, view its sheets in a browser-based grid, run the existing normalization pipeline (names, gender, dates, identifiers) through a clean service layer, manually edit cells after processing, and download the final corrected workbook — all without internet access, authentication, or a database.

The existing normalization engines (`NameEngine`, `GenderEngine`, `DateEngine`, `IdentifierEngine`), the `NormalizationOrchestrator`, the `NormalizationPipeline`, and the IO layer (`ExcelToJsonExtractor`, `ExcelWriter`) are preserved and reused. The repository is also reorganized to remove ad-hoc root-level scripts and make the overall structure cleaner and more professional.

---

## Glossary

- **WebApp**: The FastAPI-based local web application being built.
- **API**: The FastAPI HTTP endpoint layer inside `webapp/api/`.
- **ServiceLayer**: The Python service modules inside `webapp/services/` that orchestrate workbook operations and call the normalization pipeline.
- **Session**: A server-side in-memory record that tracks one user's working copy of an uploaded workbook, identified by a `session_id` UUID.
- **WorkingCopy**: A copy of the uploaded Excel file stored in the `work/` folder, on which all processing and edits are performed.
- **SourceFile**: The original uploaded Excel file stored in the `uploads/` folder, which is never modified.
- **NormalizationPipeline**: The existing `src/excel_normalization/processing/normalization_pipeline.py` class that applies all four normalization engines to JSON row data.
- **NormalizationOrchestrator**: The existing `src/excel_normalization/orchestrator.py` class that coordinates extraction, pipeline execution, and export.
- **ExcelToJsonExtractor**: The existing `src/excel_normalization/io_layer/excel_to_json_extractor.py` class that reads an Excel file into `WorkbookDataset` / `SheetDataset` structures.
- **ExportEngine**: The existing `src/excel_normalization/export/export_engine.py` class that writes the final VBA-parity export workbook.
- **SheetDataset**: The existing `src/excel_normalization/data_types.py` dataclass representing one worksheet's extracted JSON rows.
- **WorkbookDataset**: The existing `src/excel_normalization/data_types.py` dataclass representing all sheets from a workbook.
- **GridView**: The browser-side table/grid component that displays sheet data to the user.
- **OutputFolder**: The local `output/` directory where finalized export workbooks are written before download.
- **UploadsFolder**: The local `uploads/` directory where original uploaded files are stored.
- **WorkFolder**: The local `work/` directory where working copies are stored during a session.

---

## Requirements

### Requirement 1: File Upload

**User Story:** As a non-technical user, I want to upload an Excel workbook through the browser, so that I can start the normalization workflow without using the command line.

#### Acceptance Criteria

1. THE WebApp SHALL expose a `POST /api/upload` endpoint that accepts a multipart file upload of `.xlsx` or `.xlsm` files.
2. WHEN a valid Excel file is uploaded, THE WebApp SHALL save the original file to the UploadsFolder without modification and create a WorkingCopy in the WorkFolder.
3. WHEN a valid Excel file is uploaded, THE WebApp SHALL return a `session_id` UUID and the list of sheet names found in the workbook.
4. IF the uploaded file has an extension other than `.xlsx` or `.xlsm`, THEN THE WebApp SHALL return an HTTP 400 response with a user-friendly error message.
5. IF the uploaded file cannot be opened as a valid Excel workbook, THEN THE WebApp SHALL return an HTTP 422 response with a user-friendly error message.
6. THE WebApp SHALL store the SourceFile and WorkingCopy using filenames derived from the `session_id` to avoid collisions between concurrent uploads.

---

### Requirement 2: Session Management

**User Story:** As a user, I want my upload and processing state to be tracked across requests, so that I can navigate the workflow without re-uploading the file each time.

#### Acceptance Criteria

1. THE ServiceLayer SHALL maintain an in-memory session registry mapping each `session_id` to the paths of the SourceFile, WorkingCopy, and the current `WorkbookDataset`.
2. WHEN a session is created, THE ServiceLayer SHALL initialize it with the `session_id`, SourceFile path, WorkingCopy path, and a `status` of `"uploaded"`.
3. IF a request references a `session_id` that does not exist in the registry, THEN THE WebApp SHALL return an HTTP 404 response with a user-friendly error message.
4. THE WebApp SHALL NOT require a database; all session state SHALL be held in process memory.
5. WHILE the WebApp process is running, THE ServiceLayer SHALL keep all active sessions available for subsequent requests.

---

### Requirement 3: Workbook Summary

**User Story:** As a user, I want to see a summary of the uploaded workbook, so that I can understand its structure before running normalization.

#### Acceptance Criteria

1. THE WebApp SHALL expose a `GET /api/workbook/{session_id}/summary` endpoint.
2. WHEN the endpoint is called with a valid `session_id`, THE WebApp SHALL return the list of sheet names, the row count per sheet, and the detected field names per sheet.
3. THE ServiceLayer SHALL use the existing `ExcelToJsonExtractor` to extract the `WorkbookDataset` from the WorkingCopy when computing the summary.
4. IF the WorkingCopy cannot be read, THEN THE WebApp SHALL return an HTTP 500 response with a user-friendly error message.

---

### Requirement 4: Sheet Data Loading

**User Story:** As a user, I want to load and view the data from a specific sheet in a grid, so that I can inspect the records before and after normalization.

#### Acceptance Criteria

1. THE WebApp SHALL expose a `GET /api/workbook/{session_id}/sheet/{sheet_name}` endpoint.
2. WHEN the endpoint is called with a valid `session_id` and `sheet_name`, THE WebApp SHALL return the rows of that sheet as a JSON array of objects, along with the field names.
3. THE ServiceLayer SHALL retrieve the sheet data from the in-memory `WorkbookDataset` stored in the session, or re-extract it from the WorkingCopy if not yet loaded.
4. IF the requested `sheet_name` does not exist in the workbook, THEN THE WebApp SHALL return an HTTP 404 response with a user-friendly error message.
5. THE WebApp SHALL return both original field values and, after normalization has been run, the corresponding `_corrected` field values in the same row objects.

---

### Requirement 5: Normalization Execution

**User Story:** As a user, I want to run the normalization pipeline on the working copy with one click, so that the system corrects names, gender, dates, and identifiers automatically.

#### Acceptance Criteria

1. THE WebApp SHALL expose a `POST /api/workbook/{session_id}/normalize` endpoint.
2. WHEN the endpoint is called, THE ServiceLayer SHALL invoke the existing `NormalizationOrchestrator.process_workbook_json` method (or equivalent JSON-pipeline path) on the WorkingCopy.
3. THE ServiceLayer SHALL use the existing `NormalizationPipeline` with all four engines enabled: `NameEngine`, `GenderEngine`, `DateEngine`, and `IdentifierEngine`.
4. WHEN normalization completes successfully, THE ServiceLayer SHALL update the in-memory `WorkbookDataset` in the session with the normalized `SheetDataset` objects and set the session `status` to `"normalized"`.
5. WHEN normalization completes successfully, THE WebApp SHALL return a summary including the number of sheets processed, total rows processed, and per-sheet success rates.
6. IF normalization fails for a sheet, THEN THE ServiceLayer SHALL log the error, skip that sheet, and continue processing remaining sheets.
7. IF normalization fails for all sheets, THEN THE WebApp SHALL return an HTTP 500 response with a user-friendly error message.
8. THE ServiceLayer SHALL NOT modify the SourceFile during normalization; only the WorkingCopy and in-memory session state SHALL be updated.

---

### Requirement 6: Manual Cell Editing

**User Story:** As a user, I want to manually edit individual cell values after normalization, so that I can correct any remaining errors before exporting.

#### Acceptance Criteria

1. THE WebApp SHALL expose a `PATCH /api/workbook/{session_id}/sheet/{sheet_name}/cell` endpoint that accepts `row_index`, `field_name`, and `new_value` in the request body.
2. WHEN a valid edit request is received, THE ServiceLayer SHALL update the specified cell value in the in-memory `SheetDataset` for that session and sheet.
3. WHEN a valid edit request is received, THE WebApp SHALL return the updated row object confirming the change.
4. IF the `row_index` is out of range for the sheet, THEN THE WebApp SHALL return an HTTP 400 response with a user-friendly error message.
5. IF the `field_name` does not exist in the sheet's field list, THEN THE WebApp SHALL return an HTTP 400 response with a user-friendly error message.
6. THE ServiceLayer SHALL track all manual edits in the session so they are included in the final export.

---

### Requirement 7: Export and Download

**User Story:** As a user, I want to export and download the final corrected workbook as an Excel file, so that I can use the normalized data in other systems.

#### Acceptance Criteria

1. THE WebApp SHALL expose a `POST /api/workbook/{session_id}/export` endpoint.
2. WHEN the endpoint is called, THE ServiceLayer SHALL write the current in-memory `WorkbookDataset` (including all normalization results and manual edits) to a new `.xlsx` file in the OutputFolder using the existing `ExportEngine`.
3. WHEN the export file is written successfully, THE WebApp SHALL return the file as a downloadable HTTP response with `Content-Disposition: attachment` and the appropriate MIME type.
4. THE exported file SHALL be named using the original uploaded filename with a `_normalized` suffix and a timestamp to avoid collisions.
5. THE SourceFile SHALL remain unmodified after export.
6. IF the export fails, THEN THE WebApp SHALL return an HTTP 500 response with a user-friendly error message.

---

### Requirement 8: Frontend User Interface

**User Story:** As a non-technical user, I want a clean, single-page browser interface, so that I can complete the full normalization workflow without technical knowledge.

#### Acceptance Criteria

1. THE WebApp SHALL serve a single HTML page at `GET /` that provides the complete normalization workflow UI.
2. THE WebApp SHALL render the UI using server-side Jinja2 templates or a self-contained HTML/JS page served as a static file, with no external CDN dependencies, so it works fully offline.
3. THE WebApp SHALL display an upload form that allows the user to select and submit an Excel file.
4. WHEN an upload succeeds, THE WebApp SHALL display the list of available sheets and allow the user to select one to view.
5. WHEN a sheet is selected, THE WebApp SHALL display the sheet data in a scrollable grid/table showing all rows and columns.
6. THE WebApp SHALL provide a "Run Normalization" button that triggers the normalization endpoint and refreshes the grid with the updated data.
7. WHEN normalization results are displayed, THE WebApp SHALL visually distinguish original values from corrected values (e.g., by showing both in the same cell or using color coding).
8. THE WebApp SHALL allow the user to click on a cell in the grid and edit its value inline, then submit the change to the edit endpoint.
9. THE WebApp SHALL provide an "Export / Download" button that triggers the export endpoint and initiates a file download in the browser.
10. IF any API call returns an error, THE WebApp SHALL display a user-friendly error message in the UI without crashing the page.

---

### Requirement 9: Local-Only Operation

**User Story:** As a user running this on a restricted machine, I want the application to work entirely offline with no external dependencies, so that sensitive data never leaves the local computer.

#### Acceptance Criteria

1. THE WebApp SHALL run as a local FastAPI server on `localhost` with no outbound network calls.
2. THE WebApp SHALL use only local filesystem paths (`uploads/`, `work/`, `output/`) for all file storage.
3. THE WebApp SHALL NOT require a database, external cache, or any remote service.
4. THE WebApp SHALL NOT require user authentication or authorization.
5. ALL static assets (CSS, JavaScript) served by THE WebApp SHALL be bundled locally and SHALL NOT reference external CDN URLs.

---

### Requirement 10: Repository Cleanup and Restructuring

**User Story:** As a developer maintaining this project, I want the repository to be clean, well-organized, and professional, so that it is easy to understand, extend, and maintain.

#### Acceptance Criteria

1. THE WebApp entry point SHALL be located at `webapp/app.py` and SHALL be the primary user-facing way to start the application.
2. THE WebApp SHALL organize its code into `webapp/api/`, `webapp/services/`, `webapp/models/`, `webapp/templates/`, and `webapp/static/` sub-packages.
3. THE existing normalization source code SHALL remain under `src/excel_normalization/` and SHALL NOT be moved or restructured in a way that breaks existing imports or tests.
4. Root-level ad-hoc script files (e.g., `run_normalization_b.py`, `run_parity_python.py`, `run_processors_b.py`, `compare_workbooks.py`, `compare_workbooks_structural.py`, `debug_date_groups_b.py`, `debug_structural_parity.py`) SHALL be moved to a `scripts/` folder or removed if they are clearly obsolete, so the project root is clean.
5. THE project SHALL maintain a working CLI entry point (`excel-normalize` or `python -m excel_normalization.cli`) for users who prefer the command line.
6. THE `pyproject.toml` SHALL be updated to include FastAPI and Uvicorn as dependencies alongside the existing openpyxl dependency.
7. THE `README.md` SHALL be updated to document how to start the web application.

---

### Requirement 11: Error Handling and User Feedback

**User Story:** As a non-technical user, I want clear, friendly error messages when something goes wrong, so that I understand what happened and what to do next.

#### Acceptance Criteria

1. THE WebApp SHALL return all API error responses as JSON objects with a `detail` field containing a human-readable message in the language of the interface.
2. WHEN a file upload fails due to an unsupported format, THE WebApp SHALL return a message explaining which formats are accepted (`.xlsx`, `.xlsm`).
3. WHEN normalization encounters a row-level error, THE ServiceLayer SHALL log the error internally and continue processing; the user SHALL see a summary of how many rows succeeded and how many failed.
4. WHEN an export fails, THE WebApp SHALL preserve the in-memory session state so the user can retry the export without re-uploading or re-normalizing.
5. THE WebApp SHALL log all errors with sufficient detail (sheet name, row index, field name, exception message) to allow a developer to diagnose issues from the log file.

---

### Requirement 12: Integration with Existing Normalization Pipeline

**User Story:** As a developer, I want the web app to reuse the existing normalization engines without duplicating logic, so that the web and CLI paths produce identical results.

#### Acceptance Criteria

1. THE ServiceLayer SHALL instantiate `NormalizationOrchestrator` from `src/excel_normalization/orchestrator.py` to run the normalization pipeline.
2. THE ServiceLayer SHALL NOT reimplement any normalization logic; all name, gender, date, and identifier processing SHALL be delegated to the existing engines.
3. THE ServiceLayer SHALL use `ExcelToJsonExtractor` from `src/excel_normalization/io_layer/excel_to_json_extractor.py` to read workbook data into `WorkbookDataset` / `SheetDataset` structures.
4. THE ServiceLayer SHALL use `ExportEngine` from `src/excel_normalization/export/export_engine.py` to write the final export workbook.
5. FOR ALL valid Excel workbooks, running normalization via the web app endpoint SHALL produce output equivalent to running `NormalizationOrchestrator.process_workbook_json` directly (round-trip equivalence property).
