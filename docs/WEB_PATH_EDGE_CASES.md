# WEB PATH EDGE CASES — CURRENT STATE

> **Document purpose:** Exhaustive current-state reference for the web path of this project.
> Every claim is grounded in actual code. Status labels: **IMPLEMENTED** | **PARTIAL** | **MISSING** | **INACTIVE** | **INCONSISTENT**

---

## 1. Scope

### What is included

The "web path" is the browser-facing FastAPI application. It covers exactly these layers:

| Layer | Entry point | Key file |
|---|---|---|
| Upload | `POST /api/upload` | `webapp/api/upload.py` → `webapp/services/upload_service.py` |
| Sheet display | `GET /api/workbook/{id}/sheet/{name}` | `webapp/api/workbook.py` → `webapp/services/workbook_service.py` |
| standardization | `POST /api/workbook/{id}/normalize` | `webapp/api/normalize.py` → `webapp/services/standardization_service.py` → `src/excel_standardization/processing/standardization_pipeline.py` |
| Edit | `PATCH /api/workbook/{id}/sheet/{name}/cell` | `webapp/api/edit.py` → `webapp/services/edit_service.py` |
| Delete | `DELETE /api/workbook/{id}/sheet/{name}/rows` | `webapp/api/edit.py` → `webapp/services/edit_service.py` |
| Export | `POST /api/workbook/{id}/export` | `webapp/api/export.py` → `webapp/services/export_service.py` |
| Session state | (all of the above) | `webapp/services/session_service.py`, `webapp/models/session.py` |

Shared engine code called by the web path:
- `src/excel_standardization/processing/standardization_pipeline.py`
- `src/excel_standardization/engines/name_engine.py` + `text_processor.py`
- `src/excel_standardization/engines/gender_engine.py`
- `src/excel_standardization/engines/date_engine.py`
- `src/excel_standardization/engines/identifier_engine.py`
- `src/excel_standardization/io_layer/excel_to_json_extractor.py`
- `src/excel_standardization/io_layer/excel_reader.py`
- `webapp/services/derived_columns.py`
- `webapp/services/mosad_id_scanner.py`

### What is excluded

- `src/excel_standardization/orchestrator.py` — CLI / direct-Excel path only
- `src/excel_standardization/processing/*_processor.py` — processor-based Excel-writing flow, not called by web path
- `src/excel_standardization/export/export_engine.py` — used only by CLI path
- `src/excel_standardization/io_layer/excel_writer.py` — not called by web path
- Any VBA parity discussion not directly affecting web behavior

---

## 2. End-to-End Web Flow

```
1. POST /api/upload
   UploadService.handle_upload()
   - Validates extension (.xlsx / .xlsm only)
   - Saves source copy (never modified) + working copy
   - Opens workbook read-only to get sheet names
   - Creates SessionRecord(status="uploaded", workbook_dataset=None)
   - Returns: {session_id, sheet_names}

2. GET /api/workbook/{id}/sheet/{name}
   WorkbookService.get_sheet_data()
   - Lazily extracts the sheet from working_copy_path via ExcelToJsonExtractor
   - Scans for MosadID label via scan_mosad_id()
   - Stores SheetDataset in session memory
   - Applies display shaping:
       a. Strip _standardization* keys
       b. Drop completely empty rows (original columns only)
       c. Drop leading numbers-only helper row
       d. Build display_columns (original → corrected → status)
       e. apply_derived_columns() → inject _serial and MosadID
   - Returns: {sheet_name, field_names, rows}

3. POST /api/workbook/{id}/normalize[?sheet=name]
   standardizationService.normalize()
   - Re-extracts sheet(s) fresh from working_copy_path
   - Preserves MosadID from existing metadata or re-scans
   - Runs standardizationPipeline.normalize_dataset() per sheet:
       a. Detect name patterns (first 10 rows)
       b. Per row: names → gender → dates → identifiers
       c. Writes *_corrected fields + status fields into row dicts
   - Merges normalized sheets back into session dataset
   - Sets session status = "normalized"
   - Returns: {session_id, status, sheets_processed, total_rows, per_sheet_stats}

4. PATCH /api/workbook/{id}/sheet/{name}/cell
   EditService.edit_cell()
   - Validates row_index and field_name exist in in-memory row
   - Mutates sheet.rows[row_index][field_name] = new_value
   - Records edit in record.edits dict
   - Returns: {row_index, updated_row}

5. DELETE /api/workbook/{id}/sheet/{name}/rows
   EditService.delete_rows()
   - Validates all indices before touching data (all-or-nothing)
   - Removes rows in reverse order
   - Returns: {deleted_count, remaining_rows}

6. POST /api/workbook/{id}/export
   ExportService.export()
   - If workbook_dataset is None, auto-loads from disk (no standardization)
   - Applies visible_rows() filtering (same as display)
   - Maps corrected fields to fixed export schema (14 or 15 columns)
   - Creates new .xlsx workbook (RTL, right-aligned headers)
   - Returns: FileResponse with {stem}_normalized_{timestamp}.xlsx
```

---

## 3. Edge Cases by Area

### 3.1 Upload / File Validation

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| UP-01 | Wrong extension | `file.csv`, `file.xls`, `file.pdf` | HTTP 400: "File format not supported. Got: '.csv'" | **IMPLEMENTED** | Extension check before any I/O | `UploadService.handle_upload` | `suffix not in {".xlsx", ".xlsm"}` |
| UP-02 | No extension | `filename` (no dot) | HTTP 400 — `Path("filename").suffix` returns `""`, not in allowed set | **IMPLEMENTED** | Same extension check | `UploadService.handle_upload` | `suffix = ""` → rejected |
| UP-03 | Empty filename from client | `file.filename` is `None` or `""` | Defaults to `"upload.xlsx"` via `file.filename or "upload.xlsx"` in router | **IMPLEMENTED** | Router-level fallback | `webapp/api/upload.py` line: `file.filename or "upload.xlsx"` | Filename becomes `"upload.xlsx"` |
| UP-04 | Corrupt / non-Excel binary | Valid extension but corrupt bytes | HTTP 422: "could not be opened as a valid Excel workbook" | **IMPLEMENTED** | openpyxl raises on open; files deleted | `UploadService.handle_upload` | `load_workbook` raises → 422 |
| UP-05 | Workbook with zero sheets | Valid xlsx but no worksheets | HTTP 422 via `raise ValueError("Workbook has no sheets")` | **IMPLEMENTED** | Explicit check after `_wb.sheetnames` | `UploadService.handle_upload` | `if not sheet_names: raise ValueError` |
| UP-06 | File size limit | Very large file (e.g. 500MB) | **No size limit enforced** — FastAPI reads entire body into memory | **MISSING** | No `max_size` check anywhere in upload path | `UploadService.handle_upload` | 500MB file → OOM risk |
| UP-07 | Duplicate upload (same file twice) | User uploads same file again | New session_id created; old session unaffected; no deduplication | **IMPLEMENTED** (by design) | Each upload is independent | `SessionService.create` | Two sessions with different UUIDs |
| UP-08 | xlsm file | `.xlsm` uploaded | Accepted; saved as `.xlsm`; extracted with `data_only=True` (macros not executed); exported as `.xlsx` | **IMPLEMENTED** | Extension allowed; export always creates new Workbook() | `UploadService`, `ExportService.export` | Input: `file.xlsm` → Output: `file_normalized_*.xlsx` |
| UP-09 | Password-protected workbook | Encrypted xlsx | openpyxl raises `InvalidFileException` or similar → HTTP 422 | **IMPLEMENTED** (incidentally) | openpyxl cannot open encrypted files | `UploadService.handle_upload` | Raises on `load_workbook` |
| UP-10 | IO error saving file | Disk full during `write_bytes` | HTTP 500: "Failed to save the uploaded file" | **IMPLEMENTED** | try/except around file write | `UploadService.handle_upload` | `source_path.write_bytes(file_bytes)` fails |

---

### 3.2 Sheet Loading / Header Detection

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| SL-01 | Empty sheet | Sheet with no rows or cells | `detect_table_region` returns `None`; `extract_sheet_to_json` returns `SheetDataset(rows=[], field_names=[], skipped=True)` | **IMPLEMENTED** | `row_scores` is empty → returns `None` | `ExcelReader.detect_table_region` | `max_row=0` → no scores |
| SL-02 | No recognizable header | Sheet with data but no known keywords | `_score_header_row` returns score < 3 for all rows → `None` → sheet skipped | **IMPLEMENTED** | Score threshold = 3 | `ExcelReader.detect_table_region` | Sheet with only numeric data |
| SL-03 | Header beyond row 30 | Header at row 35 | Sheet skipped — `max_scan_rows=30` hard limit | **IMPLEMENTED** | `max_scan_rows=30` in `ExcelToJsonExtractor.__init__` | `ExcelToJsonExtractor.__init__` | Header at row 35 → not found |
| SL-04 | Merged header cells | `A1:C1` merged with "שם פרטי" | Value read from top-left cell of merge; all spanned columns marked as processed | **IMPLEMENTED** | `_is_merged_cell` + `_get_merged_cell_range` | `ExcelReader.detect_columns` | Merged "תאריך לידה" → year/month/day sub-columns detected |
| SL-05 | Two-row header (date groups) | Row 1: "תאריך לידה" merged; Row 2: "שנה", "חודש", "יום" | `header_rows=2`; `detect_date_groups` maps year/month/day columns | **IMPLEMENTED** | `_score_subheader_row` + `detect_date_groups` | `ExcelReader.detect_table_region`, `detect_date_groups` | `birth_year`, `birth_month`, `birth_day` fields created |
| SL-06 | Column-index helper row | Row after header: `1, 2, 3, 4, 5` (sequential integers) | `_is_column_index_row` returns `True`; `data_start_row` incremented by 1 | **IMPLEMENTED** | Requires ≥3 values, all consecutive, all ≤ end_col | `ExcelReader._is_column_index_row` | Row `[1,2,3,4,5]` → skipped |
| SL-07 | Column-index row with gaps | Row: `1, 3, 5` (non-consecutive) | `_is_column_index_row` returns `False`; row treated as data | **IMPLEMENTED** | Gap check: `sorted_vals[i] - sorted_vals[i-1] > 1` | `ExcelReader._is_column_index_row` | Row `[1,3,5]` → kept as data |
| SL-08 | Column-index row with < 3 values | Row: `1, 2` | `_is_column_index_row` returns `False` | **IMPLEMENTED** | `len(values) < 3` → False | `ExcelReader._is_column_index_row` | Row `[1,2]` → kept |
| SL-09 | Formula cells | Cell contains `=SUM(A1:A5)` | `data_only=True` → openpyxl returns computed value; if unevaluated (starts with `=`) → stored as `None` | **IMPLEMENTED** | `handle_formulas=True`; `cell_value.startswith('=')` → `None` | `ExcelToJsonExtractor.extract_row_to_json` | `=SUM(...)` unevaluated → `None` |
| SL-10 | Formula error cells | Cell contains `#VALUE!`, `#REF!` | Stored as `None` with warning logged | **IMPLEMENTED** | `cell_value.startswith('#')` → `None` | `ExcelToJsonExtractor.extract_row_to_json` | `#VALUE!` → `None` |
| SL-11 | Sheet extraction error | Unexpected exception during extraction | Returns `SheetDataset(rows=[], skipped=True, error=str(e))` | **IMPLEMENTED** | Outer try/except in `extract_sheet_to_json` | `ExcelToJsonExtractor.extract_sheet_to_json` | Any unhandled exception |
| SL-12 | Sheet not found in workbook | `GET /sheet/NonExistent` | HTTP 404: "Sheet 'NonExistent' not found" | **IMPLEMENTED** | Checked in `_ensure_sheet_loaded` | `WorkbookService._ensure_sheet_loaded` | Sheet name typo |
| SL-13 | MosadID label scanning | Sheet has "מספר מוסד: 12345" | `scan_mosad_id` finds label, reads adjacent cell value, stores in `sheet.metadata["MosadID"]` | **IMPLEMENTED** | Scans every cell; stops at first match | `webapp/services/mosad_id_scanner.py` | Label at (2,1), value at (2,2) → "12345" |
| SL-14 | Multiple MosadID labels | Two "מספר מוסד" labels in sheet | First match wins; second ignored | **IMPLEMENTED** | `scan_mosad_id` returns on first match | `scan_mosad_id` | Two labels → first value used |
| SL-15 | MosadID label with no adjacent value | Label cell exists but neighbors are empty | Returns `None`; no MosadID metadata set | **IMPLEMENTED** | `_coerce_value` returns `None` for empty | `scan_mosad_id` | Label at col 1, col 0 and col 2 empty → None |
| SL-16 | Passthrough columns (unknown headers) | Column header not in `FIELD_KEYWORDS` | Sanitized header text used as field name; column included in extraction | **IMPLEMENTED** | Passthrough pass in `detect_columns` | `ExcelReader.detect_columns` | "כתובת" → field key `"כתובת"` |

---

### 3.3 Row Filtering / Display Shaping

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| RF-01 | Completely empty row | All original-column cells are None or whitespace | Row filtered out — `any(v is not None and str(v).strip() != "")` is False | **IMPLEMENTED** | Checked against `original_field_set` only | `WorkbookService.get_sheet_data` | Row `{first_name: None, last_name: None}` → removed |
| RF-02 | Whitespace-only row | All cells contain `"   "` | Filtered — `str(v).strip() != ""` is False | **IMPLEMENTED** | Same filter as RF-01 | `WorkbookService.get_sheet_data` | Row `{first_name: "  "}` → removed |
| RF-03 | Row with corrected values but empty source | After standardization: `first_name=None`, `first_name_corrected="יוסי"` | Row is **filtered out** — filter checks only original columns | **INCONSISTENT** | Filter uses `original_field_set`; corrected fields ignored | `WorkbookService.get_sheet_data` | Normalized row with empty source → invisible in UI |
| RF-04 | Leading numbers-only helper row | First data row: `{col1: 1, col2: 2, col3: 3}` | Removed from display — `_is_numeric_like` all True | **IMPLEMENTED** | Checks only `clean_rows[0]` | `WorkbookService.get_sheet_data` | Row `[1,2,3,4,5]` → hidden |
| RF-05 | Second numbers-only row | Second row is also all-numeric | **Not removed** — check only applies to `clean_rows[0]` | **PARTIAL** | Only first row checked | `WorkbookService.get_sheet_data` | Row 2 `[1,2,3]` → shown |
| RF-06 | Display column ordering | After standardization | Original field → corrected field → status field (anchored to rightmost group member) | **IMPLEMENTED** | `_anchor_to_status` logic | `WorkbookService.get_sheet_data` | `birth_day` → `birth_day_corrected` → `birth_date_status` |
| RF-07 | Status column with no anchor | `identifier_status` exists but neither `id_number` nor `passport` in `field_names` | Status appended at end of `display_columns` via "remaining keys" loop | **IMPLEMENTED** | Falls through to `all_row_keys` loop | `WorkbookService.get_sheet_data` | Orphaned status → last column |
| RF-08 | Corrected field with no source | `first_name_corrected` in rows but `first_name` not in `field_names` | Appended at end of `display_columns` | **IMPLEMENTED** | Falls through to remaining keys | `WorkbookService.get_sheet_data` | Orphaned corrected field → last column |
| RF-09 | `_standardization_failures` key | Row has `_standardization_failures: ["gender"]` | Stripped from response — `k.startswith("_standardization")` | **IMPLEMENTED** | Metadata key filter | `WorkbookService.get_sheet_data` | Internal key never sent to client |
| RF-10 | Export applies same row filter | Export before standardization | `visible_rows()` applies identical empty-row and helper-row filters | **IMPLEMENTED** | `ExportService.visible_rows` mirrors `WorkbookService` logic | `ExportService.visible_rows` | Same rows hidden in UI and export |

---

### 3.4 Derived Columns

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| DC-01 | No serial column in source | Sheet has no "מספר סידורי" or equivalent | Synthetic `_serial` column injected with values 1, 2, 3... | **IMPLEMENTED** | `detect_serial_field` returns `None` → inject `SYNTHETIC_SERIAL_KEY` | `derived_columns.apply_derived_columns` | `_serial: 1, 2, 3` prepended |
| DC-02 | Serial column exists but some cells blank | Source has "מספר סידורי" with gaps | Blank cells auto-filled with 1-based position | **IMPLEMENTED** | `if v is None or str(v).strip() == "": row[serial_field] = i` | `derived_columns.apply_derived_columns` | Row 3 blank → filled with `3` |
| DC-03 | MosadID from metadata | `scan_mosad_id` found "12345" | Injected into every row that lacks it; shown as second column | **IMPLEMENTED** | `if not row.get(MOSAD_ID_KEY): row[MOSAD_ID_KEY] = meta_mosad_id` | `derived_columns.apply_derived_columns` | All rows get `MosadID: "12345"` |
| DC-04 | MosadID absent entirely | No label found in sheet | `meta_mosad_id=None`; MosadID column not shown (condition: `mosad_id_has_value`) | **IMPLEMENTED** | `mosad_id_has_value = any(row.get(MOSAD_ID_KEY) not in (None, ""))` | `derived_columns.apply_derived_columns` | No MosadID column in display |
| DC-05 | SugMosad column | Export schema includes "SugMosad" | `EXPORT_MAPPING["SugMosad"] = "SugMosad"` — reads from row dict; if absent → blank cell | **IMPLEMENTED** (but never populated) | No code populates `SugMosad` in web path | `ExportService.EXPORT_MAPPING` | SugMosad always blank in export |
| DC-06 | Serial column detection false positive | Field named `id_number` | `"מספר"` is in `_SERIAL_EXACT_ONLY` — must match entire key; `id_number` normalizes to `"id number"` ≠ `"מספר"` → no false positive | **IMPLEMENTED** | Exact-only set prevents substring match | `derived_columns.detect_serial_field` | `id_number` not mistaken for serial |

---

### 3.5 Name standardization

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| NM-01 | None value | `first_name=None` | `safe_to_string(None)` → `""`; `first_name_corrected=""` | **IMPLEMENTED** | Pipeline: `if original is None or original == "": json_row["first_name_corrected"] = original` | `standardizationPipeline.apply_name_standardization` | `None` → `""` |
| NM-02 | Empty string | `first_name=""` | `first_name_corrected=""` (same path as None) | **IMPLEMENTED** | Same early-return | `standardizationPipeline.apply_name_standardization` | `""` → `""` |
| NM-03 | Whitespace only | `first_name="   "` | `clean_name("   ")` → `split()=[]` → `""` | **IMPLEMENTED** | Step 5 of `clean_name` pipeline | `TextProcessor.clean_name` | `"   "` → `""` |
| NM-04 | Zero-width characters | `first_name="\u200b\u200c"` | Stripped in step 1; result `""` | **IMPLEMENTED** | `_ZERO_WIDTH` set filter | `TextProcessor.clean_name` | `"\u200b"` → `""` |
| NM-05 | Digits only | `first_name="12345"` | Language=MIXED; digits dropped; `""` | **IMPLEMENTED** | Character filter drops digits | `TextProcessor.clean_name` | `"12345"` → `""` |
| NM-06 | Symbols only | `first_name="@#$%"` | Language=MIXED; all dropped; `""` | **IMPLEMENTED** | Character filter | `TextProcessor.clean_name` | `"@#$%"` → `""` |
| NM-07 | Digits + Hebrew | `first_name="יוסי123"` | Language=HEBREW; digits dropped; `"יוסי"` | **IMPLEMENTED** | Hebrew dominant; digits not in Hebrew range | `TextProcessor.clean_name` | `"יוסי123"` → `"יוסי"` |
| NM-08 | Digits + English | `first_name="John123"` | Language=ENGLISH; digits dropped; `"John"` | **IMPLEMENTED** | English dominant | `TextProcessor.clean_name` | `"John123"` → `"John"` |
| NM-09 | Punctuation only | `first_name=".,;:!?"` | Language=MIXED; all dropped; `""` | **IMPLEMENTED** | Character filter | `TextProcessor.clean_name` | `".,;:!?"` → `""` |
| NM-10 | Hyphen variants | `"בן-דוד"`, `"בן–דוד"` (en-dash) | All `_HYPHEN_CHARS` → space; `"בן דוד"` | **IMPLEMENTED** | 8-character hyphen set | `TextProcessor.clean_name` | `"בן-דוד"` → `"בן דוד"` |
| NM-11 | Geresh / gershayim | `"ז\"ל"` | Punctuation dropped → `"זל"` → `remove_unwanted_tokens` → `""` | **IMPLEMENTED** | `HEBREW_UNWANTED_TOKENS` contains `"זל"` | `TextProcessor.clean_name` | `"ז\"ל"` → `""` |
| NM-12 | Parentheses | `"(יוסי)"` | Dropped; `"יוסי"` | **IMPLEMENTED** | Character filter | `TextProcessor.clean_name` | `"(יוסי)"` → `"יוסי"` |
| NM-13 | Hebrew diacritics (nikud) | `"יוֹסֵף"` | Nikud (U+05B0–U+05C7) outside Hebrew letter range 1488–1514; dropped | **IMPLEMENTED** | Range check `HEBREW_START <= code <= HEBREW_END` | `TextProcessor.clean_name` | `"יוֹסֵף"` → `"יסף"` (base letters only) |
| NM-14 | Title only | `"ד\"ר"` | → `"דר"` → `remove_unwanted_tokens` removes `"דר"` → `""` | **IMPLEMENTED** | `HEBREW_UNWANTED_TOKENS` contains `"דר"` | `TextProcessor.clean_name` | `"ד\"ר"` → `""` |
| NM-15 | Hebrew + English mixed | `"יוסי John"` | Hebrew count ≥ English → HEBREW; English letters dropped; `"יוסי"` | **IMPLEMENTED** | `detect_language_dominance`: Hebrew wins on tie | `TextProcessor.detect_language_dominance` | `"יוסי John"` → `"יוסי"` |
| NM-16 | Equal Hebrew/English count | `"ab יב"` (2+2) | `hebrew_count >= english_count` → HEBREW; English dropped | **IMPLEMENTED** | Tie-breaking rule: Hebrew wins | `TextProcessor.detect_language_dominance` | `"ab יב"` → `"יב"` |
| NM-17 | Single-token first name = last name | `first_name="כהן"`, `last_name="כהן"` | `len(first_name.split()) == 1` → no modification; returns `"כהן"` | **IMPLEMENTED** | Single-word guard in `remove_last_name_from_first_name` | `NameEngine.remove_last_name_from_first_name` | `"כהן"` stays `"כהן"` |
| NM-18 | Father name = last name (single token) | `father_name="כהן"`, `last_name="כהן"` | Stage A: `remove_substring("כהן","כהן")` → `""`; returns `""` | **IMPLEMENTED** | No single-word guard for father name | `NameEngine.remove_last_name_from_father` | `"כהן"` → `""` |
| NM-19 | Pattern detection sample size | Dataset with 10+ rows | Samples first 10 rows for pattern detection | **IMPLEMENTED** | `corrected_dataset.rows[:10]` | `standardizationPipeline.normalize_dataset` | 10-row sample |
| NM-20 | Pattern detection with < 3 matches | Only 2 rows have last name in first name | `contain < 3` → `FatherNamePattern.NONE`; no removal | **IMPLEMENTED** | Threshold = 3 | `NameEngine.detect_father_name_pattern` | 2 matches → NONE |
| NM-21 | Pattern applied to all rows | Pattern detected from first 10 rows | Same pattern applied to every row in dataset | **IMPLEMENTED** | `_first_name_pattern` / `_father_name_pattern` cached on pipeline instance | `standardizationPipeline.normalize_dataset` | Row 500 uses pattern from rows 1-10 |
| NM-22 | Name standardization engine failure | Engine raises unexpected exception | Fallback: `first_name_corrected = first_name` (original preserved); field added to `_standardization_failures` | **IMPLEMENTED** | try/except in `apply_name_standardization` | `standardizationPipeline.apply_name_standardization` | Exception → original value kept |


---

### 3.6 Gender standardization

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| GN-01 | None | `gender=None` | Pipeline early-return: `gender_corrected = None` (original preserved, engine not called) | **IMPLEMENTED** | `if original is None or original == "": json_row["gender_corrected"] = original` | `standardizationPipeline.apply_gender_standardization` | `None` → `gender_corrected=None` |
| GN-02 | Empty string | `gender=""` | Same early-return; `gender_corrected=""` | **IMPLEMENTED** | Same condition | `standardizationPipeline.apply_gender_standardization` | `""` → `gender_corrected=""` |
| GN-03 | Whitespace only | `gender="   "` | Engine called; `str(value).strip().lower()=""` → returns `1` | **INCONSISTENT** | Whitespace is not caught by pipeline early-return (only `None` and `""` are); engine treats it as empty → male | `GenderEngine.normalize_gender` | `"   "` → `gender_corrected=1` (not preserved as-is) |
| GN-04 | "ז" (male Hebrew) | `gender="ז"` | Not in `FEMALE_PATTERNS` → returns `1` | **IMPLEMENTED** | Pattern set check | `GenderEngine.normalize_gender` | `"ז"` → `1` |
| GN-05 | "נ" (female Hebrew) | `gender="נ"` | In `FEMALE_PATTERNS` → returns `2` | **IMPLEMENTED** | Pattern set check | `GenderEngine.normalize_gender` | `"נ"` → `2` |
| GN-06 | "f" / "F" | `gender="F"` | `lower()` → `"f"` in `FEMALE_PATTERNS` → `2` | **IMPLEMENTED** | Case-insensitive via `lower()` | `GenderEngine.normalize_gender` | `"F"` → `2` |
| GN-07 | Substring trap: "נ" in "נקבה" | `gender="נקבה"` | `"נ" in "נקבה"` = True → `2` | **IMPLEMENTED** | Substring match, not exact | `GenderEngine.normalize_gender` | `"נקבה"` → `2` |
| GN-08 | Combined value | `gender="זכר/נקבה"` | `"נ"` is substring → `2` (female wins) | **IMPLEMENTED** | First matching pattern wins | `GenderEngine.normalize_gender` | `"זכר/נקבה"` → `2` |
| GN-09 | Unknown value | `gender="unknown"` | No pattern matches → `1` (male default) | **IMPLEMENTED** | Default return | `GenderEngine.normalize_gender` | `"unknown"` → `1` |
| GN-10 | Gender engine failure | Engine raises exception | Fallback: `gender_corrected = original`; added to `_standardization_failures` | **IMPLEMENTED** | try/except in `apply_gender_standardization` | `standardizationPipeline.apply_gender_standardization` | Exception → original kept |

---

### 3.7 Date standardization

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| DT-01 | Split date — all three present | `birth_year=1990, birth_month=5, birth_day=15` | `_has_split_date` → True; `parse_from_split_columns` → valid | **IMPLEMENTED** | All three non-None/non-empty | `DateEngine.parse_from_split_columns` | → `birth_year_corrected=1990` etc. |
| DT-02 | Split date — one component missing | `birth_year=1990, birth_month=None, birth_day=15` | `_has_split_date` → False (month is None); falls to `parse_from_main_value(None)` → `status="תא ריק"` | **IMPLEMENTED** | `_has_split_date` requires all three non-None/non-empty | `DateEngine._has_split_date` | month=None → treated as no date |
| DT-03 | Split date — unparseable component | `birth_year="abc"` | `int(float("abc"))` raises → `status="תוכן לא ניתן לפריקה"` | **IMPLEMENTED** | try/except in `parse_from_split_columns` | `DateEngine.parse_from_split_columns` | `"abc"` → error status |
| DT-04 | Invalid day | `birth_day=35` | `dy > 31` → `status="יום לא תקין"`, `is_valid=False` | **IMPLEMENTED** | `_validate_date` range check | `DateEngine._validate_date` | `35` → invalid |
| DT-05 | Invalid month | `birth_month=13` | `mo > 12` → `status="חודש לא תקין"` | **IMPLEMENTED** | `_validate_date` range check | `DateEngine._validate_date` | `13` → invalid |
| DT-06 | Impossible date | `birth_year=1990, birth_month=2, birth_day=30` | `datetime(1990,2,30)` raises `ValueError` → `status="תאריך לא קיים"` | **IMPLEMENTED** | try/except around `datetime()` | `DateEngine._validate_date` | Feb 30 → invalid |
| DT-07 | Invalid components still written | `birth_month=13` | `result.year=1990, result.month=13, result.day=5` stored even when invalid; pipeline writes these to `*_corrected` | **IMPLEMENTED** | `_validate_date` always stores components; pipeline: `result.year if result.year is not None else year_val` | `standardizationPipeline._normalize_date_field` | `birth_month_corrected=13` written |
| DT-08 | 4-digit year string | `main_val="1990"` | `1900 <= 1990 <= 2100` → `year=1990, month=0, day=0, status="חסר חודש ויום"` | **IMPLEMENTED** | Special case in `_parse_numeric_date_string` | `DateEngine._parse_numeric_date_string` | `"1990"` → year only |
| DT-09 | 5 or 7 digit string | `main_val="12345"` | `len != 4,6,8` → `status="אורך תאריך לא תקין"` | **IMPLEMENTED** | Length check | `DateEngine._parse_numeric_date_string` | `"12345"` → error |
| DT-10 | 6-digit string | `main_val="150590"` | `dy=15, mo=05, yr=expand(90)` → `_validate_date` | **IMPLEMENTED** | DDMMYY format | `DateEngine._parse_numeric_date_string` | `"150590"` → 15/05/1990 |
| DT-11 | 8-digit string | `main_val="15051990"` | `dy=15, mo=05, yr=1990` → `_validate_date` | **IMPLEMENTED** | DDMMYYYY format | `DateEngine._parse_numeric_date_string` | `"15051990"` → 15/05/1990 |
| DT-12 | Excel serial integer | `raw_value=36526` | `1 <= 36526 <= 2958465` → `from_excel(36526)` → date object | **IMPLEMENTED** | Integer range check before string parsing | `DateEngine.parse_date_value` | `36526` → 2000-01-01 |
| DT-13 | Zero serial | `raw_value=0` | `1 <= 0` is False → falls to string parsing → `"0"` → `status="אורך תאריך לא תקין"` | **IMPLEMENTED** | Range check excludes 0 | `DateEngine.parse_date_value` | `0` → error |
| DT-14 | ISO-like string | `main_val="1997-09-04T00:00:00"` | Regex `^(\d{4})-(\d{2})-(\d{2})` matches → `yr=1997, mo=9, dy=4` | **IMPLEMENTED** | ISO regex before separator check | `DateEngine.parse_date_value` | `"1997-09-04T00:00:00"` → valid |
| DT-15 | Slash-separated | `main_val="15/05/1990"` | `"/" in txt` → `_parse_separated_date_string(DDMM)` | **IMPLEMENTED** | Separator detection | `DateEngine.parse_date_value` | `"15/05/1990"` → 15/05/1990 |
| DT-16 | Dot-separated | `main_val="15.05.1990"` | `"." in txt` → `replace(".","/")` → `_parse_separated_date_string` | **IMPLEMENTED** | Dot normalized to slash | `DateEngine.parse_date_value` | `"15.05.1990"` → 15/05/1990 |
| DT-17 | Two-part date (no year) | `main_val="15/05"` | `len(parts)==2` → current year injected | **IMPLEMENTED** | `parts = [parts[0], parts[1], str(date.today().year)]` | `DateEngine._parse_separated_date_string` | `"15/05"` → 15/05/2026 |
| DT-18 | English month name | `main_val="15 January 2005"` | `_contains_month_name` → True; `_parse_mixed_month_numeric` → `month=1, day=15, year=2005` | **IMPLEMENTED** | Month name dictionary | `DateEngine._parse_mixed_month_numeric` | `"15 January 2005"` → valid |
| DT-19 | Hebrew month name | `main_val="15 ינואר 2005"` | `_extract_month_number("ינואר")=1` → valid | **IMPLEMENTED** | Hebrew month dictionary | `DateEngine._parse_mixed_month_numeric` | `"15 ינואר 2005"` → valid |
| DT-20 | Two-digit year | `main_val="15/05/90"` | `yr=90 < 100` → `_expand_two_digit_year(90)` → 1990 (if 90 > current_two) | **IMPLEMENTED** | Pivot: `yr <= current_two` → current century; else previous | `DateEngine._expand_two_digit_year` | `"90"` → 1990 (in 2026) |
| DT-21 | Year before 1900 | `birth_year=1850` | `validate_business_rules`: `year < 1900` → `is_valid=False`, `status="שנה לפני 1900"` | **IMPLEMENTED** | Business rule check | `DateEngine.validate_business_rules` | `1850` → invalid |
| DT-22 | Future birth date | `birth_date > today` | `date_val > today` → `status="תאריך לידה עתידי"` | **IMPLEMENTED** | Business rule | `DateEngine.validate_business_rules` | Tomorrow → invalid |
| DT-23 | Future entry date | `entry_date > today` | `status="תאריך כניסה עתידי"` | **IMPLEMENTED** | Business rule | `DateEngine.validate_business_rules` | Tomorrow → invalid |
| DT-24 | Age over 100 | `birth_year=1900` | `age > 100` → `is_valid` stays True but `status_text="גיל מעל 100 (N שנים)"` | **IMPLEMENTED** | Age warning, not error | `DateEngine.validate_business_rules` | 1900 → warning status |
| DT-25 | Empty entry date | `entry_date=None` or `entry_date=""` | `validate_business_rules` with `ENTRY_DATE`: if `status_text=="תא ריק"` → clears to `""`, `is_valid=False` | **IMPLEMENTED** | Entry date empty is acceptable | `DateEngine.validate_business_rules` | `None` → no status written |
| DT-26 | DDMM hardcoded in web path | All date parsing | `DateFormatPattern.DDMM` always passed; no auto-detection of MMDD | **INCONSISTENT** | Pipeline always uses `DateFormatPattern.DDMM` | `standardizationPipeline._normalize_date_field` | US-format `"01/15/1990"` → parsed as 01/15 (invalid month 15) |
| DT-27 | Entry before birth — web path | `entry_date < birth_date` | **Not checked** — `DateEngine.validate_entry_before_birth` exists but is never called by pipeline | **INACTIVE** | Method exists in `DateEngine` but pipeline does not call it | `DateEngine.validate_entry_before_birth` | entry=1990, birth=2000 → no warning |
| DT-28 | datetime object in year column | `birth_year=datetime(1990,5,15)` | `isinstance(year_val, datetime)` → treated as `main_val`; parsed as date object | **IMPLEMENTED** | Special case in `_normalize_date_field` | `standardizationPipeline._normalize_date_field` | datetime → valid date |
| DT-29 | Single date field (not split) | Sheet has `birth_date` column (not year/month/day) | `has_single=True`; parsed via `parse_from_main_value`; result formatted as `"DD/MM/YYYY"` if valid | **IMPLEMENTED** | `has_single = date_field in json_row` | `standardizationPipeline._normalize_date_field` | `"15/05/1990"` → `birth_date_corrected="15/05/1990"` |
| DT-30 | Date engine failure | Engine raises unexpected exception | Fallback: original values written to `*_corrected`; `*_date_status=""` | **IMPLEMENTED** | try/except in `_normalize_date_field` | `standardizationPipeline._normalize_date_field` | Exception → originals preserved |

---

### 3.8 Identifier standardization

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| ID-01 | Both missing | `id_number=None`, `passport=None` | Early return: `id_number_corrected=None`, `passport_corrected=None`; no `identifier_status` written | **IMPLEMENTED** | `if (id_value is None or id_value == "") and (passport_value is None or passport_value == "")` | `standardizationPipeline.apply_identifier_standardization` | Both None → no status |
| ID-02 | Neither field in row | Row has no `id_number` or `passport` key | Early return with no changes | **IMPLEMENTED** | `if "id_number" not in json_row and "passport" not in json_row: return` | `standardizationPipeline.apply_identifier_standardization` | No identifier fields → skipped |
| ID-03 | Passport only | `id_number=None`, `passport="AB123456"` | `id_str=""` → `status="דרכון הוזן"` | **IMPLEMENTED** | Engine logic | `IdentifierEngine.normalize_identifiers` | `passport="AB123456"` → status set |
| ID-04 | Sentinel 9999 | `id_number="9999"` | `id_str="9999"` → `id_str=""` → treated as no ID | **IMPLEMENTED** | Explicit sentinel check | `IdentifierEngine.normalize_identifiers` | `"9999"` → `corrected_id=""` |
| ID-05 | ID with letters | `id_number="12A456789"` | Non-digit, non-dash char → `moved_to_passport=True`; if passport empty, ID moved there | **IMPLEMENTED** | Character scan | `IdentifierEngine._process_id_value` | `"12A456789"` → moved to passport |
| ID-06 | ID with space | `id_number="123 456789"` | Space is non-digit, non-dash → moved to passport | **IMPLEMENTED** | Same character scan | `IdentifierEngine._process_id_value` | `"123 456789"` → moved |
| ID-07 | ID with dot | `id_number="123.456789"` | Dot is non-digit, non-dash → moved to passport | **IMPLEMENTED** | Same | `IdentifierEngine._process_id_value` | `"123.456789"` → moved |
| ID-08 | ID with ASCII hyphen | `id_number="123-456789"` | `ord("-")=45` in `DASH_CHARS`; allowed; digits extracted → 9 digits → checksum | **IMPLEMENTED** | DASH_CHARS set | `IdentifierEngine._process_id_value` | `"123-456789"` → `"123456789"` → checksum |
| ID-09 | ID with unicode dash | `id_number="123\u2013456789"` | `ord(en-dash)=8211` in `DASH_CHARS`; allowed | **IMPLEMENTED** | DASH_CHARS includes 8211 | `IdentifierEngine._process_id_value` | en-dash → allowed |
| ID-10 | ID too short (<4 digits) | `id_number="123"` | `digit_count < 4` → moved to passport | **IMPLEMENTED** | Length check | `IdentifierEngine._process_id_value` | `"123"` → moved |
| ID-11 | ID too long (>9 digits) | `id_number="1234567890"` | `digit_count > 9` → moved to passport | **IMPLEMENTED** | Length check | `IdentifierEngine._process_id_value` | `"1234567890"` → moved |
| ID-12 | All zeros | `id_number="000000000"` | `all(ch=="0")` → `return "", False, passport, False`; not moved; `status="ת.ז. לא תקינה"` | **IMPLEMENTED** | All-zeros check | `IdentifierEngine._process_id_value` | `"000000000"` → invalid, not moved |
| ID-13 | All identical digits | `id_number="111111111"` | `len(set(padded))==1` → invalid; not moved | **IMPLEMENTED** | Identical-digit check | `IdentifierEngine._process_id_value` | `"111111111"` → invalid |
| ID-14 | Float from Excel | `id_number=123456789.0` | `_safe_to_string(123456789.0)` → `"123456789.0"`; dot is non-digit → moved to passport | **IMPLEMENTED** | str() of float includes dot | `IdentifierEngine._process_id_value` | `123456789.0` → moved to passport |
| ID-15 | Valid checksum | `id_number="039337423"` | `validate_israeli_id` → True; `status="ת.ז. תקינה"`; original string returned | **IMPLEMENTED** | Luhn-like algorithm | `IdentifierEngine.validate_israeli_id` | `"039337423"` → valid |
| ID-16 | Invalid checksum | `id_number="123456789"` | `validate_israeli_id` → False; `status="ת.ז. לא תקינה"`; padded digits returned | **IMPLEMENTED** | Checksum fails | `IdentifierEngine.validate_israeli_id` | `"123456789"` → invalid |
| ID-17 | 4-digit ID padded | `id_number="1234"` | `pad_id("1234")` → `"000001234"`; checksum on padded | **IMPLEMENTED** | `zfill(9)` | `IdentifierEngine.pad_id` | `"1234"` → `"000001234"` |
| ID-18 | Passport with spaces | `passport="AB 123 456"` | `clean_passport`: space not in allowed chars → dropped; `"AB123456"` | **IMPLEMENTED** | Character whitelist | `IdentifierEngine.clean_passport` | `"AB 123 456"` → `"AB123456"` |
| ID-19 | Passport with Hebrew | `passport="אב123"` | `1488 <= ord("א") <= 1514` → kept | **IMPLEMENTED** | Hebrew range check | `IdentifierEngine.clean_passport` | `"אב123"` → `"אב123"` |
| ID-20 | Identifier engine failure | Engine raises exception | Fallback: originals written to `*_corrected`; `identifier_status=""`; fields in `_standardization_failures` | **IMPLEMENTED** | try/except | `standardizationPipeline.apply_identifier_standardization` | Exception → originals kept |


---

### 3.9 Edit Behavior

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| ED-01 | Edit original field | `PATCH /cell` with `field_name="first_name"` | In-memory row updated; `record.edits[(sheet, idx, "first_name")] = new_value`; `first_name_corrected` NOT updated | **IMPLEMENTED** | Direct dict mutation; no re-standardization | `EditService.edit_cell` | `first_name="יוסי"` → `"יוסף"`; corrected unchanged |
| ED-02 | Edit corrected field | `PATCH /cell` with `field_name="first_name_corrected"` | Allowed if key exists in row; updates in-memory; recorded in edits | **IMPLEMENTED** | Field existence check: `if req.field_name not in row` | `EditService.edit_cell` | `first_name_corrected="יוסי"` → `"יוסף"` |
| ED-03 | Edit status field | `PATCH /cell` with `field_name="birth_date_status"` | Allowed if key exists in row | **IMPLEMENTED** | Same field existence check | `EditService.edit_cell` | Status field editable |
| ED-04 | Invalid row index | `row_index=-1` or `row_index >= len(rows)` | HTTP 400: "Row index N is out of range" | **IMPLEMENTED** | Bounds check | `EditService.edit_cell` | `-1` → 400 |
| ED-05 | Invalid field name | `field_name="nonexistent"` | HTTP 400: "Field 'nonexistent' does not exist" with available fields listed | **IMPLEMENTED** | `if req.field_name not in row` | `EditService.edit_cell` | Unknown field → 400 |
| ED-06 | new_value type | `new_value` is always `str` | Pydantic model: `new_value: str` — all edits are strings regardless of original type | **INCONSISTENT** | `CellEditRequest.new_value: str` | `webapp/models/requests.py` | Editing `birth_year` (int) → stored as `"1990"` (str) |
| ED-07 | Edit before standardization | Edit on raw (un-normalized) row | Allowed; edits raw field values | **IMPLEMENTED** | No status check | `EditService.edit_cell` | Edit `first_name` before normalize |
| ED-08 | Workbook not loaded | Edit before any sheet access | HTTP 500: "Workbook data is not available" | **IMPLEMENTED** | `if record.workbook_dataset is None` | `EditService.edit_cell` | Edit before GET /sheet → 500 |

---

### 3.10 Delete Behavior

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| DL-01 | Delete single row | `row_indices=[2]` | Row removed from `sheet.rows`; `remaining_rows` returned | **IMPLEMENTED** | `sheet.rows.pop(idx)` in reverse | `EditService.delete_rows` | Row 2 deleted |
| DL-02 | Delete multiple rows | `row_indices=[0, 3, 7]` | All validated first; deleted in reverse order | **IMPLEMENTED** | `reversed(unique_indices)` | `EditService.delete_rows` | Rows 0,3,7 deleted |
| DL-03 | Duplicate indices | `row_indices=[2, 2, 5]` | Deduplicated: `sorted(set([2,2,5]))=[2,5]`; 2 rows deleted | **IMPLEMENTED** | `unique_indices = sorted(set(req.row_indices))` | `EditService.delete_rows` | `[2,2,5]` → deletes 2 rows |
| DL-04 | One invalid index | `row_indices=[1, 999]` | HTTP 400: "Row indices out of range: [999]"; **no rows deleted** | **IMPLEMENTED** | All-or-nothing: validate all before any deletion | `EditService.delete_rows` | `[1,999]` → 400, row 1 untouched |
| DL-05 | Empty list | `row_indices=[]` | HTTP 400: "row_indices must not be empty" | **IMPLEMENTED** | `if not req.row_indices` | `EditService.delete_rows` | `[]` → 400 |
| DL-06 | Delete all rows | `row_indices=[0..n-1]` | All rows removed; `sheet.rows=[]`; `remaining_rows=0` | **IMPLEMENTED** | No minimum row count enforced | `EditService.delete_rows` | All deleted → empty sheet |

---

### 3.11 Re-Normalize Behavior

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| RN-01 | Re-normalize after edit | Edit cell → `POST /normalize` | Normalize re-extracts from `working_copy_path` (disk); **manual edits are lost** | **INCONSISTENT** | `standardizationService` always re-extracts fresh from disk | `standardizationService.normalize` | Edit `first_name` → normalize → edit gone |
| RN-02 | edits dict not replayed | `record.edits` has entries | `record.edits` is stored but **never read back** after standardization | **INACTIVE** | No code reads `record.edits` to replay | `SessionRecord.edits` | Edits recorded but silently discarded |
| RN-03 | Re-normalize after delete | Delete rows → `POST /normalize` | Deleted rows **return** — re-extraction from disk restores them | **INCONSISTENT** | Same re-extraction issue | `standardizationService.normalize` | Delete row 5 → normalize → row 5 back |
| RN-04 | Single-sheet normalize | `POST /normalize?sheet=שם` | Only that sheet re-extracted and normalized; other sheets untouched | **IMPLEMENTED** | `if sheet_name is not None:` fast path | `standardizationService.normalize` | Only "דיירים" normalized |
| RN-05 | Full normalize | `POST /normalize` (no sheet param) | All sheets re-extracted and normalized | **IMPLEMENTED** | `else:` full path | `standardizationService.normalize` | All sheets normalized |
| RN-06 | Normalize before any sheet loaded | `POST /normalize` with `workbook_dataset=None` | Auto-loads all sheets via `extract_workbook_to_json` | **IMPLEMENTED** | `if record.workbook_dataset is None:` | `standardizationService.normalize` | Normalize without prior GET /sheet |
| RN-07 | MosadID preservation on re-normalize | Re-normalize after MosadID was scanned | Preserved from existing metadata OR re-scanned: `existing.get_metadata("MosadID") or scan_mosad_id(ws)` | **IMPLEMENTED** | Explicit preservation logic | `standardizationService.normalize` | MosadID not lost on re-normalize |
| RN-08 | All sheets fail standardization | Every sheet raises exception | HTTP 500: "standardization failed for all sheets: ..." | **IMPLEMENTED** | `if not normalized_sheets: raise HTTPException(500)` | `standardizationService.normalize` | All fail → 500 |
| RN-09 | Some sheets fail | 1 of 3 sheets fails | Failed sheet skipped; others succeed; response includes only successful sheets | **IMPLEMENTED** | Per-sheet try/except; `failed_sheets` list | `standardizationService.normalize` | 2/3 succeed → response has 2 |

---

### 3.12 Export Behavior

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| EX-01 | Export before standardization | `POST /export` without prior `POST /normalize` | Auto-loads from disk (no standardization); `*_corrected` fields absent; all personal data columns blank | **IMPLEMENTED** | `if record.workbook_dataset is None:` auto-load | `ExportService.export` | Export raw → all name/date/id columns blank |
| EX-02 | No fallback to original fields | `first_name_corrected` absent | `_cell_value` returns `None`; cell left blank; no fallback to `first_name` | **IMPLEMENTED** (by design) | `EXPORT_MAPPING` maps only `*_corrected` keys | `ExportService._cell_value` | Missing corrected → blank cell |
| EX-03 | Unknown sheet name | Sheet not matching any pattern | `canonical_sheet_name` returns original name; `headers_for_sheet` returns `_HEADERS_DEFAULT` (DayarimYahidim 14-col schema) | **IMPLEMENTED** | Fallback to default schema | `ExportService.canonical_sheet_name`, `headers_for_sheet` | "Summary" → 14-col schema |
| EX-04 | xlsm input → xlsx output | `.xlsm` uploaded | Export always creates new `Workbook()` → `.xlsx` regardless of input | **IMPLEMENTED** | `ExportService` creates fresh workbook | `ExportService.export` | `file.xlsm` → `file_normalized_*.xlsx` |
| EX-05 | No highlighting in export | Changed cells | Export creates new workbook; no pink/yellow highlighting | **IMPLEMENTED** (by design) | No `ExcelWriter` used in web export | `ExportService.export` | All cells same color |
| EX-06 | RTL sheet direction | All exported sheets | `ws.sheet_view.rightToLeft = True` | **IMPLEMENTED** | Explicit setting | `ExportService.export` | All sheets RTL |
| EX-07 | SugMosad always blank | Export includes SugMosad column | `EXPORT_MAPPING["SugMosad"] = "SugMosad"`; no code populates `SugMosad` in rows | **MISSING** | No web-path code sets `SugMosad` | `ExportService.EXPORT_MAPPING` | SugMosad column always empty |
| EX-08 | MisparDiraBeMosad always blank | MeshkeyBayt/AnasheyTzevet sheets | Same issue — `MisparDiraBeMosad` never populated in web path | **MISSING** | No web-path code sets this field | `ExportService.EXPORT_MAPPING` | Column always empty |
| EX-09 | Deleted rows absent from export | Delete rows → export | `visible_rows()` uses in-memory `sheet.rows`; deleted rows absent | **IMPLEMENTED** | Export reads in-memory dataset | `ExportService.visible_rows` | Deleted row not in export |
| EX-10 | Deleted rows return after re-normalize | Delete → normalize → export | Re-normalize restores rows from disk; export includes them | **INCONSISTENT** | Re-normalize re-extracts from disk | `standardizationService.normalize` | Row deleted, then normalized → back in export |
| EX-11 | Bulk export — one session fails | `POST /export/bulk` with mixed valid/invalid sessions | Failed session skipped with warning; others exported | **IMPLEMENTED** | try/except per session in `export_bulk` | `webapp/api/export.py` | 1 of 3 fails → 2 in ZIP |
| EX-12 | Bulk export — all fail | All sessions invalid | HTTP 500: "All exports failed" | **IMPLEMENTED** | `if exported == 0: raise HTTPException(500)` | `webapp/api/export.py` | All fail → 500 |
| EX-13 | Bulk export — empty list | `session_ids=[]` | HTTP 400: "session_ids must not be empty" | **IMPLEMENTED** | `if not req.session_ids` | `webapp/api/export.py` | `[]` → 400 |
| EX-14 | Hebrew filename in Content-Disposition | Original filename has Hebrew | RFC 5987 dual-value header: ASCII fallback + `filename*=UTF-8''...` | **IMPLEMENTED** | `_content_disposition()` | `webapp/api/export.py` | `"קובץ.xlsx"` → proper header |
| EX-15 | Export output accumulates on disk | Multiple exports | Each export creates new timestamped file; old files never deleted | **MISSING** | No cleanup logic | `ExportService.export` | `output/` dir grows indefinitely |

---

### 3.13 Session / State Behavior

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location | Example |
|---|---|---|---|---|---|---|---|
| SS-01 | Session not found | Any request with unknown `session_id` | HTTP 404: "Session 'X' not found. Please upload a file first." | **IMPLEMENTED** | `_registry.get(session_id)` returns None → 404 | `SessionService.get` | Unknown UUID → 404 |
| SS-02 | Session persistence | Server restart | All sessions lost — in-memory dict only | **IMPLEMENTED** (by design) | `_registry` is module-level dict | `webapp/services/session_service.py` | Restart → all sessions gone |
| SS-03 | Concurrent requests | Two requests on same session simultaneously | No locking — single-threaded Uvicorn; race condition theoretically possible but unlikely in practice | **PARTIAL** | Comment: "No locking needed for single-threaded Uvicorn" | `SessionService` | Concurrent normalize + edit → undefined order |
| SS-04 | Session status field | `record.status` | Set to `"uploaded"` on create; set to `"normalized"` after normalize; never read by any service logic | **PARTIAL** | Status tracked but not enforced | `SessionRecord.status`, `standardizationService.normalize` | Can export without standardizing |
| SS-05 | edits dict grows unbounded | Many edits on same session | `record.edits` dict grows; never pruned; never replayed | **INACTIVE** | Dict stored but unused after recording | `SessionRecord.edits` | 1000 edits → dict has 1000 entries, all ignored |
| SS-06 | Working copy never modified | Web path | `working_copy_path` is read-only in web path; standardization re-extracts from it each time | **IMPLEMENTED** | No write operations on working copy | `standardizationService.normalize` | Working copy always original |


---

### 3.14 Error Handling

| ID | Edge Case | Input / Trigger | Current Behavior | Status | Why | Location |
|---|---|---|---|---|---|---|
| EH-01 | Session not found | Any endpoint | HTTP 404 | **IMPLEMENTED** | `SessionService.get` raises | All services |
| EH-02 | Sheet not found | `GET /sheet/X` | HTTP 404 | **IMPLEMENTED** | `_ensure_sheet_loaded` | `WorkbookService` |
| EH-03 | Workbook dataset None on edit | Edit before sheet load | HTTP 500 | **IMPLEMENTED** | `if record.workbook_dataset is None` | `EditService` |
| EH-04 | standardization total failure | All sheets fail | HTTP 500 | **IMPLEMENTED** | `if not normalized_sheets` | `standardizationService` |
| EH-05 | Export failure | Exception during workbook write | HTTP 500; session state preserved | **IMPLEMENTED** | try/except in `ExportService.export` | `ExportService` |
| EH-06 | Per-row standardization failure | Engine exception on one row | Row's original values preserved; `_standardization_failures` key added; processing continues | **IMPLEMENTED** | Per-engine try/except | `standardizationPipeline` |
| EH-07 | Per-sheet standardization failure | Sheet-level exception | Sheet skipped; others continue; if all fail → HTTP 500 | **IMPLEMENTED** | Per-sheet try/except | `standardizationService` |
| EH-08 | No structured error model used | API errors | `HTTPException` with `detail` string; `ErrorResponse` model defined but not used | **PARTIAL** | `ErrorResponse` in `responses.py` never referenced | `webapp/models/responses.py` |

---

### 3.15 Path-Specific Web Limitations

| ID | Limitation | Description | Status | Location |
|---|---|---|---|---|
| WL-01 | No DDMM/MMDD auto-detection | Date format always assumed DDMM | **MISSING** | `standardizationPipeline._normalize_date_field` |
| WL-02 | No entry-before-birth check | `DateEngine.validate_entry_before_birth` exists but not called | **INACTIVE** | `DateEngine.validate_entry_before_birth` |
| WL-03 | No cell highlighting | Export has no visual diff indicators | By design | `ExportService.export` |
| WL-04 | No edit replay after normalize | `record.edits` never replayed | **INACTIVE** | `SessionRecord.edits` |
| WL-05 | No file size limit | Large uploads accepted | **MISSING** | `UploadService.handle_upload` |
| WL-06 | No output file cleanup | Export files accumulate | **MISSING** | `ExportService.export` |
| WL-07 | new_value always string | Edit API forces string type | **INCONSISTENT** | `CellEditRequest.new_value: str` |
| WL-08 | SugMosad / MisparDiraBeMosad never populated | Export columns always blank | **MISSING** | `ExportService.EXPORT_MAPPING` |

---

## 4. Missing / Weak Areas Worth Adding

### 4.1 Entry-Before-Birth Cross-Validation

**What is missing:** The web path never checks whether `entry_date < birth_date`.

**Why it matters:** A person cannot enter an institution before being born. This is a meaningful data quality check.

**Where to add:** `standardizationPipeline.apply_date_standardization` — after both birth and entry dates are normalized, call `DateEngine.validate_entry_before_birth(birth_result, entry_result)` and write a warning to `entry_date_status`.

**Underlying helper:** `DateEngine.validate_entry_before_birth` — fully implemented, never called by pipeline.

```python
# DateEngine.validate_entry_before_birth already exists:
def validate_entry_before_birth(self, birth: DateParseResult, entry: DateParseResult) -> bool:
    # Returns False if entry < birth
```

---

### 4.2 Edit Replay After Re-Normalize

**What is missing:** `record.edits` is populated on every `PATCH /cell` call but is never read back. After `POST /normalize`, all manual edits are silently discarded.

**Why it matters:** Users who manually correct a cell and then normalize (e.g., to normalize a different sheet) lose their corrections without warning.

**Where to add:** `standardizationService.normalize` — after merging normalized sheets, iterate `record.edits` and re-apply each edit to the corresponding row.

**Underlying helper:** `record.edits` dict with keys `(sheet_name, row_idx, field_name)` already exists in `SessionRecord`.

---

### 4.3 Date Format Auto-Detection (DDMM vs MMDD)

**What is missing:** `standardizationPipeline._normalize_date_field` always passes `DateFormatPattern.DDMM`. If a sheet uses US-format dates (MM/DD/YYYY), all dates will be parsed incorrectly.

**Why it matters:** A date like `"01/15/1990"` is valid MMDD but will fail as DDMM (month=15 → invalid).

**Where to add:** `standardizationPipeline._normalize_date_field` — sample the first few non-null date values and call `DateFieldProcessor.detect_date_format_pattern` logic (or equivalent) to determine the pattern before processing all rows.

**Underlying helper:** `DateFieldProcessor.detect_date_format_pattern` exists in the direct-Excel path and could be extracted to a shared utility.

---

### 4.4 File Size Limit on Upload

**What is missing:** No maximum file size check in `UploadService.handle_upload`. The entire file is read into memory via `await file.read()`.

**Why it matters:** A 500MB file would be read entirely into memory before any validation.

**Where to add:** `webapp/api/upload.py` — FastAPI supports `max_size` on `UploadFile`, or check `len(file_bytes)` after read.

---

### 4.5 SugMosad and MisparDiraBeMosad Population

**What is missing:** `EXPORT_MAPPING` includes `"SugMosad"` and `"MisparDiraBeMosad"` but no web-path code ever sets these keys in row dicts.

**Why it matters:** The export schema reserves columns for them; they are always blank.

**Where to add:** Either during MosadID scanning (`scan_mosad_id` could be extended to also find SugMosad), or as additional metadata on `SheetDataset`.

---

### 4.6 Export Output File Cleanup

**What is missing:** `ExportService.export` creates a new timestamped file on every call. No cleanup mechanism exists.

**Why it matters:** The `output/` directory grows indefinitely in long-running deployments.

**Where to add:** `ExportService.export` — delete files older than N hours, or delete the previous export for the same session.

---

### 4.7 Edit new_value Type Coercion

**What is missing:** `CellEditRequest.new_value: str` forces all edits to be strings. Editing `birth_year` (originally an int) stores `"1990"` (str), which may cause type inconsistencies downstream.

**Why it matters:** After editing a numeric field, the export may write a string where a number is expected.

**Where to add:** `EditService.edit_cell` — attempt to coerce `new_value` to the original field's type before storing.

---

### 4.8 Whitespace-Only Gender Not Preserved

**What is missing:** `apply_gender_standardization` has an early-return for `None` and `""` but not for whitespace-only strings. `"   "` reaches the engine, which strips it and returns `1` (male). The original whitespace value is not preserved.

**Why it matters:** Inconsistent with how `None` and `""` are handled (both preserved as-is).

**Where to add:** `standardizationPipeline.apply_gender_standardization` — add `or str(original).strip() == ""` to the early-return condition.

---

## 5. Inactive Code Relevant to the Web Path

| Code | File | What it does | Why inactive |
|---|---|---|---|
| `DateEngine.validate_entry_before_birth` | `src/excel_standardization/engines/date_engine.py` | Checks if entry date precedes birth date; returns False if so | Never called by `standardizationPipeline` |
| `SessionRecord.edits` | `webapp/models/session.py` | Stores manual cell edits as `{(sheet, row, field): value}` | Populated by `EditService` but never read back after standardization |
| `ErrorResponse` model | `webapp/models/responses.py` | Pydantic model for structured error responses | Defined but never used as a response model in any router |
| `TextProcessor.remove_titles` | `src/excel_standardization/engines/text_processor.py` | Removes raw-form Hebrew/English titles before char filtering | Kept for backwards-compat; `clean_name` uses `remove_unwanted_tokens` instead |
| `TextProcessor.fix_hebrew_final_letters` | `src/excel_standardization/engines/text_processor.py` | Inserts space after final Hebrew letters | Defined but never called from `clean_name` pipeline |
| `NameEngine.normalize_names` / `normalize_first_names` / `normalize_father_names` | `src/excel_standardization/engines/name_engine.py` | Batch standardization methods | Not called by `standardizationPipeline`; pipeline calls `normalize_name` per row |
| `SessionService.delete` | `webapp/services/session_service.py` | Removes a session from registry | No API endpoint calls this; sessions accumulate for process lifetime |

---

## 6. Summary Matrix

| Functional Area | Implemented in Web | Partial in Web | Missing in Web | Main Files | Main Risk / Note |
|---|---|---|---|---|---|
| Upload / file validation | Extension, corrupt file, empty workbook | — | File size limit | `upload_service.py` | OOM on very large files |
| Sheet loading | Empty sheet, no header, merged headers, formula cells, passthrough columns | — | — | `excel_to_json_extractor.py`, `excel_reader.py` | Header scan limited to 30 rows |
| Header detection | Keyword matching, 2-row headers, date groups, column-index row | — | — | `excel_reader.py` | Score threshold may miss unusual layouts |
| Row filtering / display | Empty rows, whitespace rows, helper row, display column ordering | Second helper row not removed | — | `workbook_service.py` | Corrected-only rows filtered out (RF-03) |
| Derived columns | Serial injection, MosadID injection | SugMosad/MisparDiraBeMosad never populated | SugMosad, MisparDiraBeMosad | `derived_columns.py`, `export_service.py` | Export columns always blank |
| Name standardization | All character-level cases, pattern detection, Stage A/B removal | — | — | `text_processor.py`, `name_engine.py`, `standardization_pipeline.py` | Pattern from first 10 rows applied to all |
| Gender standardization | All patterns, case-insensitive | Whitespace-only not preserved | — | `gender_engine.py`, `standardization_pipeline.py` | Whitespace → male (GN-03) |
| Date standardization | All formats, business rules, split/single | DDMM hardcoded | MMDD auto-detection, entry-before-birth check | `date_engine.py`, `standardization_pipeline.py` | US-format dates silently wrong |
| Identifier standardization | All ID/passport cases, checksum, padding | — | — | `identifier_engine.py`, `standardization_pipeline.py` | Float IDs moved to passport |
| Edit behavior | Cell edit, field validation, index validation | new_value always string | Type coercion | `edit_service.py` | Edits lost on re-normalize |
| Delete behavior | Single/multi delete, all-or-nothing, deduplication | — | — | `edit_service.py` | Deleted rows return after re-normalize |
| Re-normalize | Single/full sheet, MosadID preservation, failure handling | — | Edit replay | `standardization_service.py` | Edits and deletes not preserved |
| Export | Fixed schema, RTL, row filtering, bulk export | — | SugMosad/MisparDiraBeMosad, file cleanup | `export_service.py` | No fallback to original fields |
| Session / state | Session CRUD, 404 handling | Status field not enforced | Session cleanup, edit replay | `session_service.py`, `session.py` | Sessions accumulate in memory |
| Error handling | Per-row/sheet failure isolation, HTTP codes | ErrorResponse model unused | — | All services | Errors are string messages only |

---

## 7. Final Assessment

### Strongest areas

**Name standardization** is the most complete area. The `TextProcessor.clean_name` pipeline handles every character-level edge case (zero-width chars, diacritics, hyphen variants, Hebrew titles, language detection). The two-stage last-name removal logic is well-designed and correctly handles single-token names.

**Identifier standardization** is thorough. The `IdentifierEngine` covers all realistic ID formats, correctly handles the float-from-Excel edge case, implements the Israeli checksum algorithm, and has clear status text for every outcome.

**Row filtering and display shaping** is consistent between the UI display path and the export path — both call the same `visible_rows()` / `apply_derived_columns()` logic.

**Error isolation** is solid. Per-row and per-sheet failures are caught independently; a single bad row or sheet does not abort the entire standardization.

### Most fragile areas

**Edit/re-normalize interaction** is the most fragile area. `record.edits` is populated but never replayed. Any re-standardization silently discards all manual corrections. There is no warning to the user. This is a significant UX and data integrity gap.

**Date format detection** is hardcoded to DDMM. A workbook with US-format dates (MM/DD/YYYY) will produce silently wrong results — dates like `"01/15/1990"` will fail with "חודש לא תקין" (month 15 invalid) rather than being correctly parsed.

**Entry-before-birth validation** is fully implemented in `DateEngine` but completely inactive in the web path. The method `validate_entry_before_birth` exists and works correctly but is never called by `standardizationPipeline`.

### Highest-priority missing behaviors

1. **Edit replay after re-normalize** — `record.edits` infrastructure exists; needs ~10 lines to replay in `standardizationService.normalize`
2. **Entry-before-birth check** — `DateEngine.validate_entry_before_birth` exists; needs one call site in `standardizationPipeline.apply_date_standardization`
3. **DDMM/MMDD auto-detection** — logic exists in direct-Excel path; needs to be extracted and called from `standardizationPipeline._normalize_date_field`
4. **File size limit** — one-line guard in `UploadService.handle_upload`
5. **Export file cleanup** — prevents unbounded disk growth in production
