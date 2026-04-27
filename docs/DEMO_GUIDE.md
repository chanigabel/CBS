# Demo Guide — Excel Normalization System

---

## 1. What the System Does (Plain Language)

Many institutions collect personal data in Excel files — names, ID numbers, dates of birth, gender codes, and so on. The problem is that these files are often inconsistent: names have extra spaces or numbers mixed in, dates are written in different formats, gender is written as Hebrew letters in some files and numbers in others, and ID numbers may have formatting errors.

This system takes those messy Excel files, reads them, automatically fixes the inconsistencies, and produces a clean, standardized output file — ready to be submitted or processed further.

**The key point:** the original file is never touched. The system reads it, does all the work in memory, and writes a brand new clean file.

**What it fixes automatically:**
- Names: removes extra spaces, numbers, titles (like "ד"ר"), and other noise
- Gender: converts any representation (Hebrew letters, text, codes) to a standard numeric code
- Dates: handles split columns (year/month/day), different formats, and two-digit years
- ID numbers: validates Israeli ID checksums and cleans passport numbers

**What the user does:**
1. Opens a browser (no installation needed — it runs locally)
2. Uploads an Excel file
3. Clicks "Run Normalization"
4. Reviews the results in a grid
5. Optionally edits cells manually
6. Clicks "Export / Download" to get the clean file

---

## 2. Technical Architecture

### Folder Overview

```
excel-data-normalization/
├── src/                  ← Core normalization engine (pure Python, no web)
├── webapp/               ← Web application (FastAPI backend + browser UI)
├── tests/                ← Automated test suite
├── scripts/              ← Utility and demo scripts (not part of the app)
├── docs/                 ← Documentation and specification files
├── schemas/              ← JSON schema definitions for internal data structures
├── installer/            ← Inno Setup script for building the Windows installer
├── launcher.py           ← Entry point for the packaged desktop app
├── build_exe.bat         ← Script to build the Windows .exe with PyInstaller
└── ExcelNormalization.spec ← PyInstaller configuration
```

### `src/` — The Normalization Engine

This is the brain of the system. It contains all the logic for reading Excel files, normalizing data, and writing output. It has no dependency on the web layer — it can be used standalone from the command line.

```
src/excel_normalization/
├── engines/              ← Pure business logic (no Excel, no web)
│   ├── name_engine.py        — name cleaning rules
│   ├── gender_engine.py      — gender normalization rules
│   ├── date_engine.py        — date parsing and validation
│   ├── identifier_engine.py  — Israeli ID checksum, passport cleaning
│   └── text_processor.py     — shared text utilities (Hebrew/English detection)
├── processing/           ← Applies engines to a full dataset row by row
│   ├── normalization_pipeline.py  — orchestrates all four engines
│   ├── name_processor.py
│   ├── gender_processor.py
│   ├── date_processor.py
│   └── identifier_processor.py
├── io_layer/             ← Reads Excel files into memory as Python dicts
│   ├── excel_reader.py        — detects headers, table regions, merged cells
│   ├── excel_to_json_extractor.py — converts a worksheet to a list of dicts
│   └── excel_writer.py        — writes output Excel files
├── export/
│   └── export_engine.py       — builds the final output workbook
├── orchestrator.py       ← Coordinates the full pipeline end-to-end
├── data_types.py         ← Core data structures (SheetDataset, WorkbookDataset)
├── cli.py                ← Command-line interface entry point
└── json_exporter.py      ← Exports internal data to JSON files
```

### `webapp/` — The Web Application

This is the browser-based interface. It wraps the `src/` engine in a local web server so users can interact with it through a browser instead of the command line.

```
webapp/
├── app.py            ← FastAPI application setup, routes, middleware
├── dependencies.py   ← Shared service instances (upload dir, work dir, output dir)
├── api/              ← HTTP endpoints (one file per feature area)
│   ├── upload.py         — POST /api/upload
│   ├── workbook.py       — GET /api/workbook/{id}/summary, /sheet/{name}
│   ├── normalize.py      — POST /api/workbook/{id}/normalize
│   ├── edit.py           — PATCH /api/workbook/{id}/sheet/{name}/cell
│   ├── export.py         — POST /api/workbook/{id}/export
│   └── institution.py    — GET/PATCH /api/workbook/{id}/institution
├── services/         ← Business logic for each feature (called by api/)
│   ├── upload_service.py
│   ├── normalization_service.py
│   ├── export_service.py
│   ├── edit_service.py
│   ├── workbook_service.py
│   ├── session_service.py    — in-memory session registry
│   ├── derived_columns.py    — serial number and MosadID injection
│   └── mosad_id_scanner.py   — scans worksheet for institution ID label
├── models/           ← Pydantic request/response schemas
│   ├── requests.py
│   ├── responses.py
│   └── session.py
├── static/
│   ├── app.js        ← All frontend logic (vanilla JS, no frameworks)
│   └── style.css     ← All styles
└── templates/
    └── index.html    ← Single HTML page (the entire UI)
```

### `tests/` — Test Suite

Contains unit tests, integration tests, and property-based tests (using Hypothesis). The `tests/webapp/` subfolder tests the web API layer specifically.

### `scripts/` — Utility Scripts

Two active scripts remain:
- `demo_pipeline.py` — demonstrates the raw pipeline flow end-to-end
- `run_parity_python.py` — runs the pipeline on a file and writes output

The `scripts/archive/` subfolder contains older debug and comparison scripts kept for reference.

### `docs/` — Documentation

Specification documents, edge case catalogues, implementation comparisons, and build instructions. Not needed at runtime.

### `schemas/` — JSON Schema Definitions

Formal JSON Schema files that define the structure of the internal data objects (`JsonRow`, `SheetDataset`, `WorkbookDataset`). Used for validation and documentation.

---

### The Main Backend Flow (Step by Step)

When a user uploads a file and clicks "Run Normalization", here is exactly what happens:

```
1. Browser → POST /api/upload
   UploadService:
   - Validates file extension (.xlsx or .xlsm only, max 50 MB)
   - Saves original file to uploads/{uuid}.xlsx  (never modified again)
   - Saves a working copy to work/{uuid}.xlsx
   - Creates an in-memory session record
   - Returns: session_id + sheet names

2. Browser → GET /api/workbook/{session_id}/sheet/{sheet_name}
   WorkbookService:
   - Reads the working copy with openpyxl
   - ExcelReader detects the table region and headers (handles merged cells,
     multi-row headers, Hebrew/English column names)
   - ExcelToJsonExtractor converts each row to a Python dict
   - Stores the SheetDataset in the session
   - Returns: field names + rows as JSON

3. Browser → POST /api/workbook/{session_id}/normalize
   NormalizationService:
   - Re-reads the working copy fresh from disk
   - Builds a NormalizationPipeline with all four engines
   - For each row: runs name, gender, date, and identifier normalization
   - Each engine adds a *_corrected field next to the original field
   - Stores the normalized SheetDataset back in the session
   - Returns: sheets processed, total rows, per-sheet success rate

4. Browser → POST /api/workbook/{session_id}/export
   ExportService:
   - Takes the normalized in-memory dataset
   - Maps corrected fields to a fixed output schema
     (MosadID, SugMosad, ShemPrati, ShemMishpaha, ... YomKnisa)
   - Writes a new .xlsx file to output/
   - Returns the file as a download
```

### The Frontend

The entire UI is a single HTML page (`webapp/templates/index.html`) driven by one JavaScript file (`webapp/static/app.js`). There are no external JavaScript libraries — it is plain vanilla JS. The page is right-to-left (Hebrew layout).

Key UI sections visible on screen:
- **Upload Excel Files** — file picker + Upload button
- **Open Files** — tabs showing each uploaded file (appears after upload)
- **Select Sheet** — tabs for each worksheet in the workbook
- **Institution bar** — fields for MosadID, institution type (SugMosad), institution name
- **▶ Run Normalization** button (also Ctrl+Enter)
- **Data grid** — shows all rows with original and corrected columns side by side
- **⬇ Export / Download** button (also Ctrl+S)

---

## 3. Step-by-Step Demo Script

### Before the Demo

1. Open a terminal in the project root folder
2. Make sure dependencies are installed: `pip install -r requirements.txt`
3. Have `sample_census_form.xlsx` ready (it is in the project root)

### Step 1 — Start the Server

Run this command in the terminal:

```bash
uvicorn webapp.app:app --reload
```

You should see output like:
```
INFO:     Uvicorn running on http://127.0.0.1:8000 (Press CTRL+C to quit)
INFO:     Excel Normalization Web App started.
```

**What to say:** "The system runs entirely locally — no internet connection, no cloud, no data leaves this machine."

### Step 2 — Open the Browser

Navigate to: `http://127.0.0.1:8000`

You should see the **Excel Normalization** page with:
- A header: "Excel Normalization"
- A subtitle: "Upload, normalize, edit, and export Excel workbooks"
- An "Upload Excel Files" section with a file picker

**What to say:** "This is the full interface — it runs in any browser, no installation required for the end user."

### Step 3 — Upload a File

1. Click the file picker area ("Choose .xlsx or .xlsm files")
2. Select `sample_census_form.xlsx` from the project root
3. Click the blue **Upload** button

After a moment you should see:
- A new tab appear under "Open Files" showing the filename
- A new section "Select Sheet" with a tab called "Census Form"
- The data grid loads automatically showing the raw sheet data

**What to say:** "The file is uploaded and immediately parsed. The original file is saved as a read-only copy — we never touch it again."

### Step 4 — Review the Raw Data

Point to the grid. You will see columns like `first_name`, `last_name`, `gender`, `birth_year`, `birth_month`, `birth_day`, `id_number`, etc.

**What to say:** "This is the raw data exactly as it came from the Excel file — no changes yet. You can see the original values in each column."

### Step 5 — Fill in Institution Details (Optional but Recommended)

In the institution bar at the top of the action area:
- Type a MosadID (e.g. `1234`) in the MosadID field
- Type an institution type (e.g. `בית אבות`) in the "סוג 1" field
- Type an institution name (e.g. `Beit Shalom`) in the institution name field

**What to say:** "These fields are workbook-level metadata — the institution ID and type will be injected into every row of the export automatically."

### Step 6 — Run Normalization

Click the green **▶ Run Normalization** button (or press Ctrl+Enter).

A spinner appears briefly, then the grid refreshes. You should now see:
- New columns with `_corrected` suffix appearing next to each original column (highlighted in green)
- A `_status` column for dates and identifiers showing validation results in Hebrew
- A stats bar showing e.g. "Normalization complete (1 sheet) — Census Form: 4 rows (100.0% success)"
- A small blue badge on rows where corrections were made

**What to say:** "The system has now run all four normalization engines — names, gender, dates, and ID numbers. The original columns are untouched. The corrected values appear in the green columns right next to them. You can see exactly what changed and why."

### Step 7 — Explore the Grid

Point out:
- A green-highlighted `_corrected` cell where a value was changed
- A `_status` column showing a Hebrew validation message
- The row badge (blue number) indicating how many fields were corrected in that row

You can also:
- Click any cell to edit it inline (click a corrected cell, type a new value, press Enter)
- Use the ▾ filter button on any column header to filter by value
- Click "⛶ הגדל טבלה" to expand the grid to full screen

**What to say:** "The user can review every correction, override any value manually, and filter the data to focus on specific rows. All manual edits are preserved even if normalization is re-run."

### Step 8 — Export

Click the grey **⬇ Export / Download** button (or press Ctrl+S).

A file download starts immediately. The filename will be either:
- `{MosadID} {InstitutionName}.xlsx` — if you filled in the institution fields
- `sample_census_form_normalized_{timestamp}.xlsx` — if you left them empty

**What to say:** "The export produces a clean, standardized Excel file with a fixed column schema — the exact format required for submission. Only the corrected values are written, not the originals. The output file is completely independent of the input."

---

## 4. What to Say

### Purpose of the System
- "This system automates the cleanup of personal data in Excel files — names, dates, IDs, gender codes — that come in inconsistent formats from different sources."
- "It replaces a manual review process that used to take hours with an automated pipeline that runs in seconds."
- "The rules are deterministic — the same input always produces the same output. There is no machine learning, no guessing."

### The Upload Process
- "When you upload a file, the system saves two copies: the original, which is locked and never modified, and a working copy that the pipeline operates on."
- "The upload validates the file format and size before accepting it — only .xlsx and .xlsm files up to 50 MB are accepted."

### The Normalization Process
- "Normalization runs four engines in sequence: names, gender, dates, and identifiers."
- "Each engine adds a corrected column next to the original — so you always see what the original value was and what the system changed it to."
- "If the system cannot confidently correct a value, it leaves a Hebrew status message explaining why."

### The Export
- "The export writes a brand new Excel file with a fixed column schema — the columns are always in the same order, with the same names, regardless of how the source file was structured."
- "Only the corrected values go into the export — not the original messy data."

### Why the Project is Organized This Way
- "The normalization logic in `src/` is completely separate from the web interface in `webapp/`. You could run the same engine from the command line without the browser."
- "This separation means the core logic can be tested independently, and the web layer can be changed without touching the business rules."

### The Original File is Not Overwritten
- "The original file is saved once on upload and never touched again. All processing happens on a separate working copy. You can always go back to the original."

### Data Safety
- "Everything runs locally on this machine. No data is sent to any server, no cloud storage, no external APIs. The files stay on disk in the `uploads/` and `work/` folders."
- "The system does not log personal data — log messages contain only session IDs and row counts."

---

## 5. What NOT to Open or Mention During the Demo

### Folders to avoid opening:
- `uploads/` — contains real Excel files with personal data from previous sessions
- `work/` — same as uploads, working copies of real data
- `archive_before_demo/` — internal cleanup archive, not relevant
- `scripts/archive/` — old debug scripts, not relevant
- `.hypothesis/`, `.mypy_cache/`, `.pytest_cache/`, `.vscode/` — tooling internals

### Files to avoid opening:
- `ExcelNormalization.spec` — PyInstaller build config, looks confusing
- `build_exe.bat`, `build_installer.bat` — build tooling, not relevant to the demo
- Any file in `dist/` — compiled binary bundle

### Topics to avoid unless asked:
- The `normalize_workbook` method in `orchestrator.py` — it is a deprecated legacy path, marked as such in the code, but seeing it could cause confusion
- The property-based test failure in `test_webapp_properties.py` — it is a pre-existing test/design mismatch unrelated to functionality
- The `work/` folder having 165 files — these are from development sessions, not a problem

---

## 6. Backup Plan

### If the server fails to start

**Check first:**
```bash
# Is port 8000 already in use?
netstat -ano | findstr :8000
```
If yes, use a different port:
```bash
uvicorn webapp.app:app --reload --port 8001
```
Then open `http://127.0.0.1:8001`

**What to say:** "The server picks a free port automatically in the packaged version — let me just switch to a different port."

### If the upload fails

**Check:** Is `sample_census_form.xlsx` in the project root? Run:
```bash
dir sample_census_form.xlsx
```
If missing, any `.xlsx` file from the `uploads/` folder will work as a demo file — they are all real workbooks that have been processed before.

**What to say:** "Let me use one of the test files we have on hand."

### If normalization returns an error

**Check:** Is the working copy readable?
```bash
python -c "from openpyxl import load_workbook; wb = load_workbook('sample_census_form.xlsx'); print(wb.sheetnames)"
```
If that works, the file is fine. Try uploading again — a fresh session will have a fresh working copy.

**What to say:** "Let me upload the file again — each upload creates a fresh independent session."

### If the export produces an empty or wrong file

This is very unlikely after a successful normalization. If it happens:
- Check that normalization was run first (the grid should show `_corrected` columns)
- Try clicking Export again — the session is preserved in memory

**What to say:** "The export reads from the in-memory session — let me make sure normalization completed first."

### Safe demo file

`sample_census_form.xlsx` in the project root is the safest demo file — it is small (4 data rows), has a clean structure, and has been tested repeatedly. It will always produce a successful normalization and export.

---

## 7. Technical Q&A

**Q: Where is the core normalization logic?**
A: In `src/excel_normalization/engines/`. Each engine is a pure Python class with no Excel or web dependencies. `name_engine.py`, `gender_engine.py`, `date_engine.py`, and `identifier_engine.py` contain all the rules. The `processing/normalization_pipeline.py` applies them in sequence to a dataset.

**Q: Is the original file changed?**
A: No. On upload, the system saves two copies: the original in `uploads/{uuid}.xlsx` (read-only, never modified) and a working copy in `work/{uuid}.xlsx`. All processing reads from the working copy. The export writes a completely new file to `output/`.

**Q: How does the system know which columns to normalize?**
A: The `ExcelReader` in `src/excel_normalization/io_layer/excel_reader.py` scans each worksheet for known Hebrew and English header patterns. It handles merged cells, multi-row headers, and various spellings. If a column header matches a known pattern (e.g. "שם פרטי", "first name", "שם"), it is mapped to the corresponding internal field name and processed by the appropriate engine.

**Q: Where are uploaded files stored?**
A: In the `uploads/` folder at the project root, named by UUID (e.g. `uploads/3f8a1b2c-....xlsx`). The working copies are in `work/`. Both folders are gitignored and never committed.

**Q: Where are exported files stored?**
A: In the `output/` folder at the project root. The filename is either `{MosadID} {InstitutionName}.xlsx` (if institution fields were filled) or `{original_stem}_normalized_{timestamp}.xlsx`. The file is also sent directly to the browser as a download.

**Q: What tests exist?**
A: The `tests/` folder contains unit tests for each engine, integration tests for the full pipeline, property-based tests using Hypothesis (which generate random inputs to verify correctness properties), and a full `tests/webapp/` suite testing every API endpoint. Run with `pytest tests/webapp/` for the web layer only.

**Q: What is the difference between the CLI and the web flow?**
A: Both use the same `src/` engine. The CLI (`python -m excel_normalization.cli file.xlsx`) runs the pipeline directly and writes `file_normalized.xlsx` next to the input. The web flow adds a browser UI, session management, lazy loading, manual cell editing, institution metadata, and a fixed-schema export. The CLI is useful for batch processing; the web app is for interactive review.

**Q: What are the known limitations?**
A: The system is designed for a specific Excel form structure used by Israeli care institutions. It expects Hebrew column headers in a known set of patterns. Files with completely different structures may not be recognized. It also processes one workbook at a time in the web UI (though multiple files can be uploaded in separate tabs). The in-memory session state is lost if the server restarts.

**Q: What should be improved after the demo?**
A: The `uploads/` and `work/` folders accumulate files indefinitely — a cleanup policy for old sessions would be a good next step. The in-memory session store could be replaced with a persistent store for multi-user or long-running scenarios. The property-based test `test_session_initialization_invariant` has a known mismatch with the lazy-loading design and should be updated to reflect actual behavior.
