# Current-State Case Catalogue


> **Purpose:** Exhaustive behavioral matrix derived from the real codebase only.
> No desired behavior. No redesign. No suggestions.
> Every rule is cited to the file and function that implements it.
> Status: **IMPLEMENTED** | **PARTIAL** | **AMBIGUOUS**

---

## 1. Files Inspected

| File | Role |
|---|---|
| `src/excel_normalization/engines/text_processor.py` | `clean_name` pipeline, language detection, char filtering, token removal |
| `src/excel_normalization/engines/name_engine.py` | `normalize_name`, `remove_last_name_from_first_name`, `remove_last_name_from_father`, pattern detection |
| `src/excel_normalization/engines/gender_engine.py` | `normalize_gender` |
| `src/excel_normalization/engines/date_engine.py` | `parse_date`, all sub-parsers, `validate_business_rules` |
| `src/excel_normalization/engines/identifier_engine.py` | `normalize_identifiers`, `_process_id_value`, `validate_israeli_id`, `clean_passport` |
| `src/excel_normalization/io_layer/excel_reader.py` | `detect_table_region`, `detect_columns`, `find_header`, `_is_column_index_row` |
| `src/excel_normalization/io_layer/excel_to_json_extractor.py` | `extract_sheet_to_json`, `extract_row_to_json` |
| `src/excel_normalization/processing/normalization_pipeline.py` | `normalize_dataset`, `normalize_row`, all `apply_*` methods |
| `src/excel_normalization/processing/name_processor.py` | Direct-Excel name processing |
| `src/excel_normalization/processing/date_processor.py` | Direct-Excel date processing |
| `src/excel_normalization/orchestrator.py` | `process_worksheet`, `_validate_entry_vs_birth` |
| `src/excel_normalization/export/export_engine.py` | VBA-parity export, `detect_header_row`, `is_valid_data_row` |
| `webapp/services/workbook_service.py` | `get_sheet_data`, display_columns ordering, row filtering |
| `webapp/services/export_service.py` | `ExportService.export`, `visible_rows`, `canonical_sheet_name` |
| `webapp/services/edit_service.py` | `edit_cell`, `delete_rows` |
| `webapp/services/normalization_service.py` | `normalize`, re-extract on normalize |
| `webapp/services/derived_columns.py` | `apply_derived_columns`, `detect_serial_field` |
| `webapp/services/mosad_id_scanner.py` | `scan_mosad_id` |
| `webapp/services/upload_service.py` | `handle_upload` |
| `webapp/services/session_service.py` | Session registry |
| `webapp/static/app.js` | UI rendering, edit commit, delete, normalization trigger |

---

## 2. Global Flow

### Web / JSON path (used by the web app)

```
User uploads .xlsx/.xlsm
  → UploadService.handle_upload()
      → UUID session created; file saved to uploads/ and work/
      → openpyxl opened read-only to get sheet names only
      → SessionRecord created with workbook_dataset=None

User clicks sheet tab
  → GET /api/workbook/{id}/sheet/{name}
      → WorkbookService._ensure_sheet_loaded()
          → ExcelToJsonExtractor.extract_sheet_to_json()
              → ExcelReader.detect_table_region()   [header scoring]
              → ExcelReader.detect_columns()         [keyword matching + passthrough]
              → extract_row_to_json() per data row
          → scan_mosad_id()                          [label/value scan]
          → SheetDataset stored in session
      → WorkbookService.get_sheet_data()
          → empty-row filter (original cols only)
          → numeric helper-row filter (first row only)
          → display_columns ordering (orig → corrected → status)
          → apply_derived_columns()                  [serial + MosadID injection]
      → SheetDataResponse returned to browser

User clicks "Run Normalization"
  → POST /api/workbook/{id}/normalize
      → NormalizationService.normalize()
          → Re-extracts ALL sheets fresh from disk (discards prior in-memory state)
          → NormalizationPipeline.normalize_dataset() per sheet
              → detect patterns (first 10 rows)
              → normalize_row() per row:
                  names → gender → dates → identifiers
          → Replaces in-memory SheetDataset with normalized version

User clicks "Export / Download"
  → POST /api/workbook/{id}/export
      → ExportService.export()
          → visible_rows() per sheet (same filters as UI)
          → canonical_sheet_name() mapping
          → headers_for_sheet() schema selection
          → EXPORT_MAPPING: only *_corrected keys, no fallback
          → New .xlsx written to output/
          → FileResponse returned
```

### Direct-Excel / CLI path

```
CLI: python -m excel_normalization.cli file.xlsx
  → NormalizationOrchestrator.normalize_workbook()
      → load_workbook(data_only=False, keep_vba=is_macro)
      → process_worksheet() per sheet:
          → NameFieldProcessor.process_field()
              → find_header() [xlPart scan]
              → prepare_output_columns() [insert_cols in-place]
              → process_data() [read → normalize → write → highlight]
          → GenderFieldProcessor.process_field()
          → DateFieldProcessor.process_field()
              → _collect_date_groups() [Find/FindNext semantics]
              → detect_date_format_pattern() [DDMM vs MMDD from data]
              → insert 4 corrected columns per date group
          → IdentifierFieldProcessor.process_field()
          → _validate_entry_vs_birth() [cross-check, appends to status cell]
      → workbook.save(file_path)  [in-place modification]
```

---

## 3. Case Catalogue

---

### 3.1 Sheet Loading

**Code:** `ExcelReader.detect_table_region()` — `excel_reader.py`

**Processing order:**
1. Score each row 1–30 (keyword matches, text density, cell count)
2. Select highest-scoring row (min score 3)
3. Check if selected row is a sub-header (שנה/חודש/יום) → look for parent above
4. Check if row below is a sub-header → set header_rows=2
5. Check for merged parent headers with date keywords
6. Detect column-index row immediately after headers → skip it
7. Find table end (5 consecutive empty rows = stop)

| Case ID | Input / Condition | Code Path | Decision | Output | Status |
|---|---|---|---|---|---|
| SL-01 | Sheet with headers in row 1, data from row 2 | `detect_table_region` | score ≥ 3 | `header_rows=1`, `data_start_row=2` | IMPLEMENTED |
| SL-02 | Sheet with title rows above headers (e.g., logo in row 1, headers in row 3) | `detect_table_region` scoring | Row 3 scores highest | `start_row=3`, `data_start_row=4` | IMPLEMENTED |
| SL-03 | Sheet with two-row header (parent + שנה/חודש/יום sub-row) | `detect_table_region` sub-header check | `header_rows=2`, `data_start_row=parent+2` | Correct | IMPLEMENTED |
| SL-04 | Sheet where best row scores < 3 | `detect_table_region` | Returns `None` | Sheet skipped; `SheetDataset(skipped=True)` | IMPLEMENTED |
| SL-05 | Sheet with no cells at all | `detect_table_region` | `row_scores` empty | Returns `None` → sheet skipped | IMPLEMENTED |
| SL-06 | Sheet with column-index row (1,2,3…) after headers | `_is_column_index_row` | All ints, ≥3, consecutive, ≤ end_col | `data_start_row` incremented by 1 | IMPLEMENTED |
| SL-07 | Column-index row with a gap (1,2,4) | `_is_column_index_row` | `sorted_vals[i] - sorted_vals[i-1] > 1` → False | Row NOT skipped; treated as data | IMPLEMENTED |
| SL-08 | Column-index row with only 2 values | `_is_column_index_row` | `len(values) < 3` → False | Row NOT skipped | IMPLEMENTED |
| SL-09 | Table ends with 5 consecutive empty rows | `_find_table_end_row` | `row_idx > last_data_row + 5` → break | `end_row` = last row with data | IMPLEMENTED |
| SL-10 | Table ends with 4 consecutive empty rows then data | `_find_table_end_row` | Does not break at 4 | Continues scanning | IMPLEMENTED |
| SL-11 | Sheet with merged cells in header row | `detect_columns` merged-cell handling | Value read from top-left of merge; all spanned cols marked processed | Correct field mapping | IMPLEMENTED |
| SL-12 | Header column containing "מתוקן" or "corrected" | `_should_ignore_column` | Returns True | Column excluded from mapping | IMPLEMENTED |
| SL-13 | Header column not matching any keyword | Passthrough pass in `detect_columns` | Safe key built from raw header text | Column included with sanitised key | IMPLEMENTED |
| SL-14 | Two-row header: name fields only on sub-header row | Sub-header pass in `detect_columns` | Scans subheader_row for unmapped cols | Fields found and mapped | IMPLEMENTED |
| SL-15 | `max_scan_rows=30` default; headers in row 31 | `detect_table_region` | Row 31 not scanned | Sheet skipped | IMPLEMENTED |
| SL-16 | `.xlsm` file | `extract_workbook_to_json` | `load_workbook(data_only=True)` — VBA not executed | Data extracted normally | IMPLEMENTED |

---

### 3.2 Header Detection

**Code:** `ExcelReader.find_header()` (direct-Excel path), `ExcelReader.detect_columns()` (JSON path)

| Case ID | Input / Condition | Code Path | Decision | Output | Status |
|---|---|---|---|---|---|
| HD-01 | Cell contains "שם פרטי" | `detect_columns` keyword match | `first_name` matched | `column_mapping["first_name"]` set | IMPLEMENTED |
| HD-02 | Cell contains "שם משפחה" | keyword match | `last_name` matched | `column_mapping["last_name"]` set | IMPLEMENTED |
| HD-03 | Cell contains "שם האב" | keyword match | `father_name` matched | `column_mapping["father_name"]` set | IMPLEMENTED |
| HD-04 | Cell contains "מין" | keyword match | `gender` matched | `column_mapping["gender"]` set | IMPLEMENTED |
| HD-05 | Cell contains "מספר זהות" or "תעודת זהות" or "ת.ז" | keyword match | `id_number` matched | `column_mapping["id_number"]` set | IMPLEMENTED |
| HD-06 | Cell contains "דרכון" | keyword match | `passport` matched | `column_mapping["passport"]` set | IMPLEMENTED |
| HD-07 | Cell contains "תאריך לידה" with sub-row שנה/חודש/יום | `detect_date_groups` + `detect_columns` | `birth_year`, `birth_month`, `birth_day` mapped | Split date fields | IMPLEMENTED |
| HD-08 | Cell contains "תאריך כניסה" with sub-row שנה/חודש/יום | Same | `entry_year`, `entry_month`, `entry_day` mapped | Split date fields | IMPLEMENTED |
| HD-09 | Cell contains "תאריך לידה" but no sub-row | `detect_columns` | `birth_date` mapped as single field | Single date field | IMPLEMENTED |
| HD-10 | Multiple keywords in one cell (e.g., "שם פרטי ומשפחה") | `_match_field` longest-match | Longest matching keyword wins | One field mapped | IMPLEMENTED |
| HD-11 | Header cell is empty but part of merged range | Merged-cell resolution | Value read from top-left of merge | Correct | IMPLEMENTED |
| HD-12 | Header cell contains "שם פרטי - מתוקן" | `_should_ignore_column` | "מתוקן" found → ignored | Column excluded | IMPLEMENTED |
| HD-13 | Direct-Excel path: `find_header(["שם פרטי"])` | `find_header` xlPart scan | Substring match anywhere in cell | Returns `ColumnHeaderInfo` | IMPLEMENTED |
| HD-14 | Direct-Excel path: header cell contains "מתוקן" | `find_header` | `if "מתוקן" in cell_text: continue` | Skipped | IMPLEMENTED |
| HD-15 | Date header: direct-Excel path searches "תאריך לידה" | `DateFieldProcessor._collect_date_groups` | Scans all cells for substring match | All matching cells collected | IMPLEMENTED |
| HD-16 | Date header: direct-Excel path searches "תאריך כניסה למוסד" | Same | Exact term required | Only cells containing that exact string | IMPLEMENTED |

---

### 3.3 Row Filtering

**Code:** `WorkbookService.get_sheet_data()`, `ExportService.visible_rows()` — both paths identical

| Case ID | Input / Condition | Code Path | Decision | Output | Visible in UI | In Export |
|---|---|---|---|---|---|---|
| RF-01 | Row where all original-column cells are None | Empty-row filter | `any(v not None and strip != "")` → False | Row dropped | No | No |
| RF-02 | Row where all original-column cells are `""` | Empty-row filter | Same | Row dropped | No | No |
| RF-03 | Row where all original-column cells are whitespace-only | Empty-row filter | `str(v).strip() != ""` → False | Row dropped | No | No |
| RF-04 | Row where original cols are empty but `_corrected` cols have values | Empty-row filter | Check is against `original_field_set` only | Row dropped (corrected values ignored in check) | No | No |
| RF-05 | Row with at least one non-empty original cell | Empty-row filter | `any(...)` → True | Row kept | Yes | Yes |
| RF-06 | First row is all-numeric (e.g., 1, 2, 3, 4…) | Numeric helper-row filter | `all(_is_numeric_like(v))` → True | First row dropped | No | No |
| RF-07 | First row has mix of numeric and text | Numeric helper-row filter | `all(...)` → False | Row kept | Yes | Yes |
| RF-08 | First row is all-numeric but second row is also all-numeric | Numeric helper-row filter | Only first row checked | Only first row dropped; second kept | Second: Yes | Second: Yes |
| RF-09 | `_normalization_failures` key in row | Metadata strip | `k.startswith("_normalization")` → stripped | Key removed from display | Not shown | Not in export |
| RF-10 | `_normalization_statistics` key in metadata | Not in rows | N/A | Not visible | No | No |
| RF-11 | Row deleted via `EditService.delete_rows` | In-memory `sheet.rows.pop(idx)` | Removed from list | Gone from session | No | No |
| RF-12 | Row deleted, then re-normalize called | `NormalizationService.normalize` re-extracts from disk | Deleted row reappears | Row is back | Yes | Yes |

---

### 3.4 Column Ordering (Display)

**Code:** `WorkbookService.get_sheet_data()` — display_columns construction

**Processing order:**
1. For each original field (Excel left-to-right order): place original, then its `_corrected`, then its status anchor
2. Append any remaining unseen keys
3. `apply_derived_columns`: prepend `_serial` (or source serial col), then `MosadID` (if any row has value)

| Case ID | Input / Condition | Decision | Output Column Order | Status |
|---|---|---|---|---|
| CO-01 | Sheet with `first_name`, `last_name` only, no normalization | No `_corrected` in rows | `[_serial, first_name, last_name]` | IMPLEMENTED |
| CO-02 | Sheet after normalization, all fields present | `_corrected` keys in rows | `[_serial, MosadID, first_name, first_name_corrected, last_name, last_name_corrected, ...]` | IMPLEMENTED |
| CO-03 | Sheet with split birth date (year, month, day) | `birth_date_status` anchors to rightmost of group | `[..., birth_year, birth_year_corrected, birth_month, birth_month_corrected, birth_day, birth_day_corrected, birth_date_status]` | IMPLEMENTED |
| CO-04 | Sheet with `id_number` and `passport` | `identifier_status` anchors to rightmost of `{id_number, passport}` | `[..., id_number, id_number_corrected, passport, passport_corrected, identifier_status]` | IMPLEMENTED |
| CO-05 | Sheet with `passport` only (no `id_number`) | `identifier_status` anchors to `passport` | `[..., passport, passport_corrected, identifier_status]` | IMPLEMENTED |
| CO-06 | Sheet with `id_number` only (no `passport`) | `identifier_status` anchors to `id_number` | `[..., id_number, id_number_corrected, identifier_status]` | IMPLEMENTED |
| CO-07 | Source sheet has a serial-number column ("מספר סידורי") | `detect_serial_field` finds it | Source column used; blanks auto-filled with position | IMPLEMENTED |
| CO-08 | Source sheet has no serial-number column | `detect_serial_field` returns None | Synthetic `_serial` injected; values = 1,2,3… | IMPLEMENTED |
| CO-09 | MosadID found in sheet metadata | `apply_derived_columns` | `MosadID` column appears at position 1 | IMPLEMENTED |
| CO-10 | MosadID not found | `mosad_id_has_value` = False | `MosadID` column not shown | IMPLEMENTED |
| CO-11 | Status key present in rows but no group member in `original_fields` | `_anchor_to_status` not set | Status key appended at end as "unexpected extra" | IMPLEMENTED |
| CO-12 | `_corrected` key present in rows but original not in `field_names` | Not placed in main loop | Appended at end as "unexpected extra" | IMPLEMENTED |

---

---

### 3.5 Name Normalization

**Code:** `TextProcessor.clean_name()` — `text_processor.py`

**Fixed pipeline order:** SafeToString → zero-width strip → diacritics → Arabic-Indic digits → language detection → char filter → space normalize → unwanted token removal

#### 3.5.1 Null / Empty / Whitespace

| Case ID | Input | After Step 1 | Language | After Filter | Final Output | Status |
|---|---|---|---|---|---|---|
| NM-01 | `None` | `""` | — | — | `""` | IMPLEMENTED |
| NM-02 | `""` | `""` | — | — | `""` | IMPLEMENTED |
| NM-03 | `"   "` | `"   "` | MIXED | `"   "` → spaces only | `""` | IMPLEMENTED |
| NM-04 | `"\u200b"` (zero-width space only) | `""` after strip | — | — | `""` | IMPLEMENTED |
| NM-05 | `"\u200b יוסי"` | `" יוסי"` | HEBREW | `" יוסי"` | `"יוסי"` | IMPLEMENTED |
| NM-06 | `0` (integer) | `"0"` | MIXED | `""` (digit dropped) | `""` | IMPLEMENTED |
| NM-07 | `1.5` (float) | `"1.5"` | MIXED | `""` (digit and dot dropped) | `""` | IMPLEMENTED |

#### 3.5.2 Digits Only / Symbols Only

| Case ID | Input | Language | After Filter | Final Output | Status |
|---|---|---|---|---|---|
| NM-08 | `"123"` | MIXED (no letters) | digits dropped | `""` | IMPLEMENTED |
| NM-09 | `"123abc"` | ENGLISH (3 English letters) | digits dropped, letters kept | `"abc"` | IMPLEMENTED |
| NM-10 | `"123יוסי"` | HEBREW (4 Hebrew letters) | digits dropped, Hebrew kept | `"יוסי"` | IMPLEMENTED |
| NM-11 | `"---"` | MIXED | hyphens → spaces → collapse | `""` | IMPLEMENTED |
| NM-12 | `"!!!"` | MIXED | dropped | `""` | IMPLEMENTED |
| NM-13 | `"(יוסי)"` | HEBREW | `(` and `)` dropped | `"יוסי"` | IMPLEMENTED |
| NM-14 | `"٠١٢٣"` (Arabic-Indic digits) | MIXED after conversion to `"0123"` | digits dropped | `""` | IMPLEMENTED |

#### 3.5.3 Hebrew Only

| Case ID | Input | Language | After Filter | After Token Removal | Final Output | Status |
|---|---|---|---|---|---|---|
| NM-15 | `"יוסי"` | HEBREW | `"יוסי"` | unchanged | `"יוסי"` | IMPLEMENTED |
| NM-16 | `"יוסי כהן"` | HEBREW | `"יוסי כהן"` | unchanged | `"יוסי כהן"` | IMPLEMENTED |
| NM-17 | `"  יוסי   כהן  "` | HEBREW | `"  יוסי   כהן  "` | `"יוסי כהן"` after collapse | `"יוסי כהן"` | IMPLEMENTED |
| NM-18 | `"יוסי ז\"ל"` | HEBREW | `"יוסי זל"` (geresh dropped) | `"זל"` removed → `"יוסי"` | `"יוסי"` | IMPLEMENTED |
| NM-19 | `"ד\"ר יוסי"` | HEBREW | `"דר יוסי"` | `"דר"` removed → `"יוסי"` | `"יוסי"` | IMPLEMENTED |
| NM-20 | `"רבי יוסי"` | HEBREW | `"רבי יוסי"` | `"רבי"` removed → `"יוסי"` | `"יוסי"` | IMPLEMENTED |
| NM-21 | `"ר יוסי"` | HEBREW | `"ר יוסי"` | `"ר"` removed (whole-token) → `"יוסי"` | `"יוסי"` | IMPLEMENTED |
| NM-22 | `"שליט\"א יוסי"` | HEBREW | `"שליטא יוסי"` | `"שליטא"` removed → `"יוסי"` | `"יוסי"` | IMPLEMENTED |
| NM-23 | `"הי\"ד יוסי"` | HEBREW | `"היד יוסי"` | `"היד"` removed → `"יוסי"` | `"יוסי"` | IMPLEMENTED |
| NM-24 | `"זצ\"ל יוסי"` | HEBREW | `"זצל יוסי"` | `"זצל"` removed → `"יוסי"` | `"יוסי"` | IMPLEMENTED |
| NM-25 | `"זיע\"א יוסי"` | HEBREW | `"זיעא יוסי"` | `"זיעא"` removed → `"יוסי"` | `"יוסי"` | IMPLEMENTED |

#### 3.5.4 English Only

| Case ID | Input | Language | After Filter | After Token Removal | Final Output | Status |
|---|---|---|---|---|---|---|
| NM-26 | `"John"` | ENGLISH | `"John"` | unchanged | `"John"` | IMPLEMENTED |
| NM-27 | `"John Smith"` | ENGLISH | `"John Smith"` | unchanged | `"John Smith"` | IMPLEMENTED |
| NM-28 | `"Dr. John"` | ENGLISH | `"Dr John"` (dot dropped) | `"Dr"` → `"dr"` matched → removed | `"John"` | IMPLEMENTED |
| NM-29 | `"John Jr."` | ENGLISH | `"John Jr"` | `"Jr"` → `"jr"` matched → removed | `"John"` | IMPLEMENTED |
| NM-30 | `"Mr. Smith"` | ENGLISH | `"Mr Smith"` | `"Mr"` removed | `"Smith"` | IMPLEMENTED |
| NM-31 | `"Mrs. Smith"` | ENGLISH | `"Mrs Smith"` | `"Mrs"` removed | `"Smith"` | IMPLEMENTED |
| NM-32 | `"Prof. Smith"` | ENGLISH | `"Prof Smith"` | `"Prof"` removed | `"Smith"` | IMPLEMENTED |
| NM-33 | `"John III"` | ENGLISH | `"John III"` | `"iii"` matched (case-insensitive) → removed | `"John"` | IMPLEMENTED |
| NM-34 | `"John IV"` | ENGLISH | `"John IV"` | `"iv"` matched → removed | `"John"` | IMPLEMENTED |
| NM-35 | `"JOHN"` | ENGLISH | `"JOHN"` | unchanged | `"JOHN"` (case preserved) | IMPLEMENTED |

#### 3.5.5 Mixed Hebrew / English

| Case ID | Input | Hebrew Count | English Count | Language | After Filter | Final Output | Status |
|---|---|---|---|---|---|---|---|
| NM-36 | `"יוסי Smith"` | 4 | 5 | ENGLISH (5>4) | Hebrew dropped | `"Smith"` | IMPLEMENTED |
| NM-37 | `"יוסי Sm"` | 4 | 2 | HEBREW (4>2) | English dropped | `"יוסי"` | IMPLEMENTED |
| NM-38 | `"יוסי Jo"` | 4 | 2 | HEBREW | English dropped | `"יוסי"` | IMPLEMENTED |
| NM-39 | `"יוסי John"` | 4 | 4 | HEBREW (tie → Hebrew wins) | English dropped | `"יוסי"` | IMPLEMENTED |
| NM-40 | `"AB יב"` | 2 | 2 | HEBREW (tie) | English dropped | `"יב"` | IMPLEMENTED |
| NM-41 | `"123"` (no letters) | 0 | 0 | MIXED | digits dropped | `""` | IMPLEMENTED |
| NM-42 | `"יוסי123"` | 4 | 0 | HEBREW | digits dropped | `"יוסי"` | IMPLEMENTED |

#### 3.5.6 Punctuation / Special Characters

| Case ID | Input | What Happens | Final Output | Status |
|---|---|---|---|---|
| NM-43 | `"יוסי-כהן"` (ASCII hyphen) | Hyphen → space | `"יוסי כהן"` | IMPLEMENTED |
| NM-44 | `"יוסי–כהן"` (en-dash U+2013) | En-dash → space | `"יוסי כהן"` | IMPLEMENTED |
| NM-45 | `"יוסי—כהן"` (em-dash U+2014) | Em-dash → space | `"יוסי כהן"` | IMPLEMENTED |
| NM-46 | `"יוסי−כהן"` (minus sign U+2212) | Minus → space | `"יוסי כהן"` | IMPLEMENTED |
| NM-47 | `"יוסי\u2011כהן"` (non-breaking hyphen) | → space | `"יוסי כהן"` | IMPLEMENTED |
| NM-48 | `"ג'ורג'"` (ASCII apostrophe) | Apostrophe dropped (not in Hebrew range, not hyphen) | `"גורג"` | IMPLEMENTED |
| NM-49 | `"ג\u05f3ורג"` (geresh U+05F3) | U+05F3 is NOT in Hebrew range (05D0–05EA) → dropped | `"גורג"` | IMPLEMENTED |
| NM-50 | `"צ\u05f4ל"` (gershayim U+05F4) | U+05F4 NOT in Hebrew range → dropped | `"צל"` | IMPLEMENTED |
| NM-51 | `"יוסי, כהן"` | Comma dropped | `"יוסי כהן"` | IMPLEMENTED |
| NM-52 | `"יוסי. כהן"` | Dot dropped | `"יוסי כהן"` | IMPLEMENTED |
| NM-53 | `"(יוסי כהן)"` | Parens dropped | `"יוסי כהן"` | IMPLEMENTED |

#### 3.5.7 Diacritics

| Case ID | Input | After Diacritic Removal | Language | Final Output | Status |
|---|---|---|---|---|---|
| NM-54 | `"André"` | `"Andre"` | ENGLISH | `"Andre"` | IMPLEMENTED |
| NM-55 | `"Müller"` | `"Muller"` | ENGLISH | `"Muller"` | IMPLEMENTED |
| NM-56 | `"Ñoño"` | `"Nono"` | ENGLISH | `"Nono"` | IMPLEMENTED |
| NM-57 | `"Ç"` | `"C"` | ENGLISH | `"C"` | IMPLEMENTED |
| NM-58 | `"ё"` (Cyrillic) | `"e"` | ENGLISH | `"e"` | IMPLEMENTED |
| NM-59 | `"à la"` | `"a la"` | ENGLISH | `"a la"` | IMPLEMENTED |

#### 3.5.8 No Surviving Letters

| Case ID | Input | Language | After Filter | Final Output | Status |
|---|---|---|---|---|---|
| NM-60 | `"123!!!"` | MIXED | all dropped | `""` | IMPLEMENTED |
| NM-61 | `"ד\"ר"` (only a title) | HEBREW | `"דר"` | `"דר"` removed → `""` | IMPLEMENTED |
| NM-62 | `"ז\"ל"` | HEBREW | `"זל"` | `"זל"` removed → `""` | IMPLEMENTED |
| NM-63 | `"Dr."` | ENGLISH | `"Dr"` | `"dr"` removed → `""` | IMPLEMENTED |

---

### 3.6 Last-Name Removal — First Name

**Code:** `NameEngine.remove_last_name_from_first_name()`, `detect_first_name_pattern()` — `name_engine.py`
**Pipeline:** `NormalizationPipeline.apply_name_normalization()` — `normalization_pipeline.py`

**Pattern detection:** Samples first 10 rows with both `first_name` and `last_name` non-empty. Uses first 5 of those. Requires `contain >= 3` to activate. Uses **raw** (pre-clean) values from rows, then calls `normalize_name` on each sample value.

| Case ID | Input (cleaned first_name) | Cleaned last_name | Pattern | Stage A | Stage B | Final Output | Status |
|---|---|---|---|---|---|---|---|
| FN-01 | `"כהן יוסי"` | `"כהן"` | REMOVE_FIRST | `" כהן "` found → `"יוסי"` (changed) | not run | `"יוסי"` | IMPLEMENTED |
| FN-02 | `"יוסי כהן"` | `"כהן"` | REMOVE_LAST | `" כהן "` found → `"יוסי"` (changed) | not run | `"יוסי"` | IMPLEMENTED |
| FN-03 | `"יוסי"` (single word) | `"כהן"` | REMOVE_FIRST | single-word guard → return | not run | `"יוסי"` | IMPLEMENTED |
| FN-04 | `"לוי יוסי"` | `"כהן"` | REMOVE_FIRST | `"כהן"` not in `"לוי יוסי"` → no-op | runs: drop first token | `"יוסי"` | IMPLEMENTED |
| FN-05 | `"לוי יוסי"` | `"כהן"` | NONE | Stage A no-op | Stage B skipped (NONE) | `"לוי יוסי"` | IMPLEMENTED |
| FN-06 | `"כהן"` (single word = last name) | `"כהן"` | REMOVE_FIRST | single-word guard → return | not run | `"כהן"` | IMPLEMENTED |
| FN-07 | `"כהן כהן"` (last name twice) | `"כהן"` | REMOVE_FIRST | `" כהן "` found in `" כהן כהן "` → replaces first occurrence → `"כהן"` | not run | `"כהן"` | IMPLEMENTED |
| FN-08 | `""` (empty after clean) | `"כהן"` | any | `not first_name` → return `""` | not run | `""` | IMPLEMENTED |
| FN-09 | `"יוסי"` | `""` (empty last name) | any | `not last_name` → return `"יוסי"` | not run | `"יוסי"` | IMPLEMENTED |
| FN-10 | `"כהן"` (equals last name, single word) | `"כהן"` | REMOVE_LAST | single-word guard → return | not run | `"כהן"` | IMPLEMENTED |
| FN-11 | Pattern detection: 2 of 5 rows contain last name | `detect_first_name_pattern` | `contain=2 < 3` → NONE | N/A | N/A | NONE pattern | IMPLEMENTED |
| FN-12 | Pattern detection: 3 of 5 rows, last name at start | `detect_first_name_pattern` | `first_pos=3 >= 3` → REMOVE_FIRST | N/A | N/A | REMOVE_FIRST | IMPLEMENTED |
| FN-13 | Pattern detection: 3 of 5 rows, last name at end | `detect_first_name_pattern` | `last_pos=3 >= 3` → REMOVE_LAST | N/A | N/A | REMOVE_LAST | IMPLEMENTED |
| FN-14 | Pattern detection: 3 of 5 rows contain, but neither first nor last position ≥ 3 | `detect_first_name_pattern` | `contain>=3` but neither pos ≥ 3 → NONE | N/A | N/A | NONE | IMPLEMENTED |
| FN-15 | Stage A: `remove_substring` result is empty string | `after_stage_a.strip()` empty | Returns `""` | not run | `""` | IMPLEMENTED |
| FN-16 | `"כהנים יוסי"` (last name is substring of word, not whole word) | `"כהן"` | REMOVE_FIRST | `" כהן "` NOT in `" כהנים יוסי "` → no-op | runs: drop first token → `"יוסי"` | `"יוסי"` | IMPLEMENTED |

---

### 3.7 Last-Name Removal — Father Name

**Code:** `NameEngine.remove_last_name_from_father()`, `detect_father_name_pattern()` — `name_engine.py`

Identical two-stage logic to first name, with one difference: **NONE pattern → never modify** (explicit guard at top of function, before Stage A).

| Case ID | Input (cleaned father_name) | Cleaned last_name | Pattern | Stage A | Stage B | Final Output | Status |
|---|---|---|---|---|---|---|---|
| FA-01 | `"כהן יוסף"` | `"כהן"` | REMOVE_FIRST | `"כהן"` found → `"יוסף"` | not run | `"יוסף"` | IMPLEMENTED |
| FA-02 | `"יוסף כהן"` | `"כהן"` | REMOVE_LAST | `"כהן"` found → `"יוסף"` | not run | `"יוסף"` | IMPLEMENTED |
| FA-03 | `"יוסף"` (single word) | `"כהן"` | REMOVE_FIRST | single-word guard → return | not run | `"יוסף"` | IMPLEMENTED |
| FA-04 | `"לוי יוסף"` | `"כהן"` | REMOVE_FIRST | not substring → no-op | runs: drop first → `"יוסף"` | `"יוסף"` | IMPLEMENTED |
| FA-05 | `"לוי יוסף"` | `"כהן"` | NONE | NONE guard at top → return immediately | not run | `"לוי יוסף"` | IMPLEMENTED |
| FA-06 | `""` | `"כהן"` | REMOVE_FIRST | `not father_name` → return `""` | not run | `""` | IMPLEMENTED |
| FA-07 | `"יוסף"` | `""` | REMOVE_FIRST | `not last_name` → return `"יוסף"` | not run | `"יוסף"` | IMPLEMENTED |
| FA-08 | `"כהן"` (equals last name, single word) | `"כהן"` | REMOVE_FIRST | single-word guard → return | not run | `"כהן"` | IMPLEMENTED |
| FA-09 | Pattern: 3 of 5 rows, last name first | `detect_father_name_pattern` | `first >= 3` → REMOVE_FIRST | N/A | N/A | REMOVE_FIRST | IMPLEMENTED |
| FA-10 | Pattern: 3 of 5 rows, last name last | `detect_father_name_pattern` | `last >= 3` → REMOVE_LAST | N/A | N/A | REMOVE_LAST | IMPLEMENTED |
| FA-11 | Pattern: 0 rows with both fields | `detect_father_name_pattern` | `sample_size=0` → NONE | N/A | N/A | NONE | IMPLEMENTED |
| FA-12 | Stage A result empty | `after_stage_a.strip()` empty | Returns `""` | not run | `""` | IMPLEMENTED |

**Key difference from first-name removal:** In `remove_last_name_from_father`, the NONE check is at the very top (before Stage A). In `remove_last_name_from_first_name`, NONE only blocks Stage B; Stage A still runs if the substring is present.

---

### 3.8 Gender Normalization

**Code:** `GenderEngine.normalize_gender()` — `gender_engine.py`
**Pipeline:** `NormalizationPipeline.apply_gender_normalization()` — `normalization_pipeline.py`

**Algorithm:** `str(value).strip().lower()` → check if any female pattern is a substring → return 2 if yes, 1 if no.

**Female patterns (substring match, case-insensitive):** `"2"`, `"female"`, `"נ"`, `"אישה"`, `"בת"`, `"f"`, `"נקבה"`, `"girl"`, `"woman"`

| Case ID | Input | Lowercased | Female Pattern Match | Output | Visible in UI | In Export | Status |
|---|---|---|---|---|---|---|---|
| GN-01 | `None` | — | — | `1` | Yes | Yes | IMPLEMENTED |
| GN-02 | `""` | `""` | empty → default | `1` | Yes | Yes | IMPLEMENTED |
| GN-03 | `"   "` | `""` after strip | empty → default | `1` | Yes | Yes | IMPLEMENTED |
| GN-04 | `1` (int) | `"1"` | no match | `1` | Yes | Yes | IMPLEMENTED |
| GN-05 | `2` (int) | `"2"` | `"2"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-06 | `"1"` | `"1"` | no match | `1` | Yes | Yes | IMPLEMENTED |
| GN-07 | `"2"` | `"2"` | `"2"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-08 | `"ז"` | `"ז"` | no match | `1` | Yes | Yes | IMPLEMENTED |
| GN-09 | `"נ"` | `"נ"` | `"נ"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-10 | `"זכר"` | `"זכר"` | no match | `1` | Yes | Yes | IMPLEMENTED |
| GN-11 | `"נקבה"` | `"נקבה"` | `"נקבה"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-12 | `"אישה"` | `"אישה"` | `"אישה"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-13 | `"בת"` | `"בת"` | `"בת"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-14 | `"male"` | `"male"` | no match | `1` | Yes | Yes | IMPLEMENTED |
| GN-15 | `"female"` | `"female"` | `"female"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-16 | `"FEMALE"` | `"female"` | `"female"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-17 | `"f"` | `"f"` | `"f"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-18 | `"F"` | `"f"` | `"f"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-19 | `"girl"` | `"girl"` | `"girl"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-20 | `"woman"` | `"woman"` | `"woman"` matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-21 | `"M"` | `"m"` | no match | `1` | Yes | Yes | IMPLEMENTED |
| GN-22 | `"unknown"` | `"unknown"` | no match | `1` | Yes | Yes | IMPLEMENTED |
| GN-23 | `"נ/א"` | `"נ/א"` | `"נ"` is substring → matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-24 | `"זכר/נקבה"` | `"זכר/נקבה"` | `"נ"` is substring of `"נקבה"` → matches | `2` | Yes | Yes | IMPLEMENTED |
| GN-25 | `None` or `""` in pipeline | `apply_gender_normalization` | `original is None or == ""` → `gender_corrected = original` | `None` or `""` | Yes | Yes | IMPLEMENTED |

**Note on GN-23/GN-24:** The female pattern check is a substring match. `"נ"` will match any string containing the Hebrew letter Nun, including `"נ/א"` or `"זכר/נקבה"`. This is a known behavior.

---

---

### 3.9 Date Parsing and Statuses

**Code:** `DateEngine` — `date_engine.py`
**Pipeline (JSON path):** `NormalizationPipeline._normalize_date_field()` — always uses `DateFormatPattern.DDMM`
**Pipeline (direct-Excel path):** `DateFieldProcessor._process_date_field()` — detects DDMM vs MMDD from data

#### 3.9.1 Split vs Single Detection (JSON path)

**Code:** `NormalizationPipeline._normalize_date_field()`

| Case ID | Row Keys Present | year_val | month_val | day_val | Path Taken | Status |
|---|---|---|---|---|---|---|
| DP-01 | `birth_year`, `birth_month`, `birth_day` all present | 1980 | 5 | 15 | Split path → `parse_from_split_columns` | IMPLEMENTED |
| DP-02 | `birth_year` only (month/day absent from row) | `"15/05/1980"` | None | None | `year_val not None, month/day None` → treated as `main_val` → single-value path | IMPLEMENTED |
| DP-03 | `birth_year` is a `datetime` object | `datetime(1980,5,15)` | None | None | `isinstance(year_val, datetime)` → treated as `main_val` | IMPLEMENTED |
| DP-04 | `birth_date` key present | — | — | — | Single path → `parse_date(None,None,None, date_val, DDMM, BIRTH_DATE)` | IMPLEMENTED |
| DP-05 | Both `birth_year` and `birth_date` present | any | any | any | `has_split` checked first → split path wins | IMPLEMENTED |
| DP-06 | `birth_year` present but all three are None | None | None | None | `_has_split_date(None,None,None)` → False → falls to `parse_from_main_value(None)` | IMPLEMENTED |

#### 3.9.2 `_has_split_date` Logic

**Code:** `DateEngine._has_split_date()`

| Case ID | y | m | d | Result | Status |
|---|---|---|---|---|---|
| HS-01 | `None` | `5` | `15` | False (y is None) | IMPLEMENTED |
| HS-02 | `""` | `5` | `15` | False (y is `""`) | IMPLEMENTED |
| HS-03 | `1980` (int) | `5` (int) | `15` (int) | True (all int/float) | IMPLEMENTED |
| HS-04 | `1980.0` (float) | `5.0` | `15.0` | True (all float) | IMPLEMENTED |
| HS-05 | `"1980"` | `"5"` | `"15"` | True (all digit strings) | IMPLEMENTED |
| HS-06 | `"1980"` | `"5"` | `"abc"` | False (`"abc"` not digit) | IMPLEMENTED |
| HS-07 | `"1980-01"` | `"5"` | `"15"` | True (hyphens stripped before isdigit check) | IMPLEMENTED |

#### 3.9.3 Split Column Parsing

**Code:** `DateEngine.parse_from_split_columns()`

| Case ID | year_val | month_val | day_val | Parsed yr/mo/dy | is_valid | status_text | Status |
|---|---|---|---|---|---|---|---|
| SC-01 | `1980` | `5` | `15` | 1980/5/15 | True | `""` | IMPLEMENTED |
| SC-02 | `80` | `5` | `15` | 1980/5/15 (two-digit expand) | True | `""` | IMPLEMENTED |
| SC-03 | `27` | `5` | `15` | 1927/5/15 (as of 2026: 27 > 26) | True | `""` | IMPLEMENTED |
| SC-04 | `25` | `5` | `15` | 2025/5/15 (as of 2026: 25 ≤ 26) | True | `""` | IMPLEMENTED |
| SC-05 | `"abc"` | `5` | `15` | — | False | `"תוכן לא ניתן לפריקה"` | IMPLEMENTED |
| SC-06 | `1980` | `13` | `15` | 1980/13/15 stored | False | `"חודש לא תקין"` | IMPLEMENTED |
| SC-07 | `1980` | `5` | `32` | 1980/5/32 stored | False | `"יום לא תקין"` | IMPLEMENTED |
| SC-08 | `1980` | `2` | `30` | 1980/2/30 stored | False | `"תאריך לא קיים"` | IMPLEMENTED |
| SC-09 | `0` | `5` | `15` | 0/5/15 stored | False | `"שנה לא תקינה"` | IMPLEMENTED |
| SC-10 | `1980` | `5` | `0` | 1980/5/0 stored | False | `"יום לא תקין"` | IMPLEMENTED |

#### 3.9.4 Single-Value Parsing — Input Type Dispatch

**Code:** `DateEngine.parse_date_value()`

| Case ID | raw_value | Type | Path | Result yr/mo/dy | is_valid | status_text | Status |
|---|---|---|---|---|---|---|---|
| SV-01 | `None` | NoneType | `raw_value is None` | None/None/None | False | `"תא ריק"` | IMPLEMENTED |
| SV-02 | `""` | str | `txt == ""` | None/None/None | False | `"תא ריק"` | IMPLEMENTED |
| SV-03 | `"   "` | str | `txt.strip() == ""` | None/None/None | False | `"תא ריק"` | IMPLEMENTED |
| SV-04 | `datetime(1980,5,15)` | datetime | isinstance check | 1980/5/15 | True | `""` | IMPLEMENTED |
| SV-05 | `date(1980,5,15)` | date | isinstance check | 1980/5/15 | True | `""` | IMPLEMENTED |
| SV-06 | `29221` (int, Excel serial) | int, 1≤x≤2958465 | `from_excel(29221)` | 1980/1/1 | True | `""` | IMPLEMENTED |
| SV-07 | `0` (int) | int, 0 < 1 | Falls through to numeric string | `"0"` → 1-char → `"אורך תאריך לא תקין"` | False | `"אורך תאריך לא תקין"` | IMPLEMENTED |
| SV-08 | `2958466` (int, > max) | int, > 2958465 | Falls through to numeric string | `"2958466"` → 7 chars → `"אורך תאריך לא תקין"` | False | `"אורך תאריך לא תקין"` | IMPLEMENTED |
| SV-09 | `"January 15 1980"` | str with month name | `_contains_month_name` → True | 1980/1/15 | True | `""` | IMPLEMENTED |
| SV-10 | `"15 ינואר 1980"` | str with Hebrew month | `_contains_month_name` → True | 1980/1/15 | True | `""` | IMPLEMENTED |
| SV-11 | `"15051980"` | str, all digits, 8 chars | `txt.isdigit()` → numeric path | 1980/5/15 | True | `""` | IMPLEMENTED |
| SV-12 | `"150580"` | str, all digits, 6 chars | numeric path | 1980/5/15 | True | `""` | IMPLEMENTED |
| SV-13 | `"1980"` | str, all digits, 4 chars, 1900≤x≤2100 | numeric path | year=1980, month=0, day=0 | False | `"חסר חודש ויום"` | IMPLEMENTED |
| SV-14 | `"1234"` | str, all digits, 4 chars, not 1900–2100 | numeric path DMYY | d=1,m=2,yr=expand(34) | depends | depends | IMPLEMENTED |
| SV-15 | `"1997-09-04T00:00:00"` | str, ISO-like | regex `^\d{4}-\d{2}-\d{2}` | 1997/9/4 | True | `""` | IMPLEMENTED |
| SV-16 | `"15/05/1980"` | str with `/` | separated path, DDMM | 1980/5/15 | True | `""` | IMPLEMENTED |
| SV-17 | `"15.05.1980"` | str with `.` | `.` → `/`, separated path | 1980/5/15 | True | `""` | IMPLEMENTED |
| SV-18 | `"05/15/1980"` | str with `/`, DDMM default | separated path, DDMM | day=5, mo=15 → `"חודש לא תקין"` | False | `"חודש לא תקין"` | IMPLEMENTED |
| SV-19 | `"abc"` | str, no digits, no separator | no path matches | None/None/None | False | `"פורמט תאריך לא מזוהה"` | IMPLEMENTED |
| SV-20 | `"15/5"` (two-part) | str with `/`, 2 parts | assumes current year | 2026/5/15 (as of 2026) | True | `""` | IMPLEMENTED |

#### 3.9.5 Numeric Date String Parsing

**Code:** `DateEngine._parse_numeric_date_string()`

| Case ID | Input | Length | Interpretation | yr/mo/dy | is_valid | status_text | Status |
|---|---|---|---|---|---|---|---|
| ND-01 | `"15051980"` | 8 | DDMMYYYY | 1980/5/15 | True | `""` | IMPLEMENTED |
| ND-02 | `"150580"` | 6 | DDMMYYyy | 1980/5/15 | True | `""` | IMPLEMENTED |
| ND-03 | `"1980"` | 4, 1900≤x≤2100 | Year only | 1980/0/0 | False | `"חסר חודש ויום"` | IMPLEMENTED |
| ND-04 | `"1234"` | 4, not 1900–2100 | DMYY: d=1,m=2,yr=expand(34) | 2034/2/1 (as of 2026) | True | `""` | IMPLEMENTED |
| ND-05 | `"12345"` | 5 | `"אורך תאריך לא תקין"` | None/None/None | False | `"אורך תאריך לא תקין"` | IMPLEMENTED |
| ND-06 | `"1234567"` | 7 | `"אורך תאריך לא תקין"` | None/None/None | False | `"אורך תאריך לא תקין"` | IMPLEMENTED |
| ND-07 | `"00000000"` | 8 | DDMMYYYY: d=0,m=0,y=0 | 0/0/0 stored | False | `"יום לא תקין"` | IMPLEMENTED |
| ND-08 | `"32011980"` | 8 | d=32 | 1980/1/32 stored | False | `"יום לא תקין"` | IMPLEMENTED |
| ND-09 | `"15131980"` | 8 | m=13 | 1980/13/15 stored | False | `"חודש לא תקין"` | IMPLEMENTED |
| ND-10 | `"30021980"` | 8 | d=30,m=2 | 1980/2/30 stored | False | `"תאריך לא קיים"` | IMPLEMENTED |

#### 3.9.6 Business Rules

**Code:** `DateEngine.validate_business_rules()`

| Case ID | Parsed Result | field_type | Rule Applied | is_valid After | status_text After | Status |
|---|---|---|---|---|---|---|
| BR-01 | `is_valid=False, status="תא ריק"`, ENTRY_DATE | ENTRY_DATE | Empty entry date special case | False | `""` (cleared) | IMPLEMENTED |
| BR-02 | `is_valid=False, status="תא ריק"`, BIRTH_DATE | BIRTH_DATE | No special case | False | `"תא ריק"` (unchanged) | IMPLEMENTED |
| BR-03 | `is_valid=False` (any other status) | any | `not result.is_valid` → return early | False | unchanged | IMPLEMENTED |
| BR-04 | `is_valid=True`, year=1899 | any | `year < 1900` | False | `"שנה לפני 1900"` | IMPLEMENTED |
| BR-05 | `is_valid=True`, year=1900 | any | `year >= 1900` → passes | True | `""` | IMPLEMENTED |
| BR-06 | `is_valid=True`, date > today | BIRTH_DATE | `date_val > today` | False | `"תאריך לידה עתידי"` | IMPLEMENTED |
| BR-07 | `is_valid=True`, date > today | ENTRY_DATE | `date_val > today` | False | `"תאריך כניסה עתידי"` | IMPLEMENTED |
| BR-08 | `is_valid=True`, age = 100 exactly | BIRTH_DATE | `age > 100` → False (100 is not > 100) | True | `""` | IMPLEMENTED |
| BR-09 | `is_valid=True`, age = 101 | BIRTH_DATE | `age > 100` → True | True (is_valid NOT changed) | `"גיל מעל 100 (101 שנים)"` | IMPLEMENTED |
| BR-10 | `is_valid=True`, age = 150 | BIRTH_DATE | `age > 100` | True | `"גיל מעל 100 (150 שנים)"` | IMPLEMENTED |
| BR-11 | `is_valid=True`, entry date | ENTRY_DATE | No age check for entry dates | True | `""` | IMPLEMENTED |

#### 3.9.7 Corrected Field Writing (JSON path)

**Code:** `NormalizationPipeline._normalize_date_field()`

| Case ID | Scenario | `birth_year_corrected` | `birth_month_corrected` | `birth_day_corrected` | `birth_date_status` | Status |
|---|---|---|---|---|---|---|
| CW-01 | Split, valid date | `result.year` | `result.month` | `result.day` | `""` | IMPLEMENTED |
| CW-02 | Split, invalid (e.g., bad month) | `result.year` (stored even if invalid) | `result.month` | `result.day` | `"חודש לא תקין"` | IMPLEMENTED |
| CW-03 | Split, completely unparseable | `year_val` (original) | `month_val` (original) | `day_val` (original) | `"תוכן לא ניתן לפריקה"` | IMPLEMENTED |
| CW-04 | Single, valid date | — | — | — | `"DD/MM/YYYY"` formatted | `""` | IMPLEMENTED |
| CW-05 | Single, parsed but invalid (e.g., Feb 30) | — | — | — | `"30/02/2000"` formatted | `"תאריך לא קיים"` | IMPLEMENTED |
| CW-06 | Single, completely unparseable | — | — | — | original value | `"פורמט תאריך לא מזוהה"` | IMPLEMENTED |
| CW-07 | Single, `None` or `""` | — | — | — | `None` or `""` (no status written) | `""` (no status key) | IMPLEMENTED |

**Note on CW-07:** When `date_val is None or == ""`, the pipeline returns early before calling the engine. `birth_date_corrected` is set to the original value, but `birth_date_status` is NOT written to the row. The status key will be absent from the row dict.

#### 3.9.8 Entry-Before-Birth Cross-Validation

| Case ID | Path | Behavior | Status |
|---|---|---|---|
| EB-01 | Direct-Excel path | `_validate_entry_vs_birth()` called after both date groups processed; appends `"תאריך כניסה לפני תאריך לידה"` to entry status cell; formats pink+bold | IMPLEMENTED |
| EB-02 | JSON / web path | `validate_entry_before_birth()` exists in `DateEngine` but is NOT called by `NormalizationPipeline` | NOT IMPLEMENTED |
| EB-03 | Direct-Excel: birth or entry not valid | `not birth.is_valid or not entry.is_valid` → returns True (no warning) | IMPLEMENTED |
| EB-04 | Direct-Excel: entry == birth date | `entry_date < birth_date` → False → no warning | IMPLEMENTED |

---

---

### 3.10 Identifier / ID / Passport Logic

**Code:** `IdentifierEngine.normalize_identifiers()`, `_process_id_value()`, `validate_israeli_id()`, `clean_passport()` — `identifier_engine.py`
**Pipeline:** `NormalizationPipeline.apply_identifier_normalization()` — `normalization_pipeline.py`

#### 3.10.1 Pre-processing and Entry Conditions

| Case ID | id_value | passport_value | Pre-processing | id_str after | passport after clean | Status |
|---|---|---|---|---|---|---|
| ID-01 | `None` | `None` | `_safe_to_string` | `""` | `""` | IMPLEMENTED |
| ID-02 | `""` | `""` | strip | `""` | `""` | IMPLEMENTED |
| ID-03 | `"9999"` | `None` | sentinel check | `""` (treated as no ID) | `""` | IMPLEMENTED |
| ID-04 | `"9999"` | `"AB123"` | sentinel check | `""` | `"AB123"` | IMPLEMENTED |
| ID-05 | `123456789` (int) | `None` | `str(123456789)` | `"123456789"` | `""` | IMPLEMENTED |
| ID-06 | `123456789.0` (float) | `None` | `str(123456789.0)` = `"123456789.0"` | `"123456789.0"` | `""` | AMBIGUOUS — float ID contains `.` which is not digit/dash → moved to passport |

#### 3.10.2 No ID Cases

| Case ID | id_str | cleaned_passport | corrected_id | corrected_passport | status_text | Status |
|---|---|---|---|---|---|---|
| NI-01 | `""` | `""` | `""` | `""` | `"חסר מזהים"` | IMPLEMENTED |
| NI-02 | `""` | `"AB123"` | `""` | `"AB123"` | `"דרכון הוזן"` | IMPLEMENTED |
| NI-03 | `"9999"` | `""` | `""` | `""` | `"חסר מזהים"` | IMPLEMENTED |
| NI-04 | `"9999"` | `"AB123"` | `""` | `"AB123"` | `"דרכון הוזן"` | IMPLEMENTED |

#### 3.10.3 ID Character Scan — Move to Passport

| Case ID | id_str | passport before | Non-digit/non-dash char? | Move? | corrected_id | corrected_passport | status_text | Status |
|---|---|---|---|---|---|---|---|---|
| MC-01 | `"AB123456"` | `""` | `A` at pos 0 | Yes (passport empty) | `""` | `"AB123456"` (cleaned) | `"ת.ז. הועברה לדרכון"` | IMPLEMENTED |
| MC-02 | `"AB123456"` | `"XY789"` | `A` at pos 0 | No (passport not empty) | `""` | `"XY789"` (unchanged) | `"ת.ז. הועברה לדרכון"` | IMPLEMENTED |
| MC-03 | `"123 456"` | `""` | space at pos 3 | Yes | `""` | `"123456"` (space dropped by clean_passport) | `"ת.ז. הועברה לדרכון"` | IMPLEMENTED |
| MC-04 | `"123.456"` | `""` | `.` at pos 3 | Yes | `""` | `"123456"` | `"ת.ז. הועברה לדרכון"` | IMPLEMENTED |
| MC-05 | `"123-456"` | `""` | `-` is in DASH_CHARS | No non-digit/non-dash | Not moved | digits=`"123456"` → 6 digits → pad → validate | depends | IMPLEMENTED |
| MC-06 | `"123–456"` (en-dash) | `""` | en-dash ord=8211 in DASH_CHARS | Not moved | digits=`"123456"` | depends | IMPLEMENTED |

#### 3.10.4 All-Zeros / All-Identical Rejection

| Case ID | id_str | digits | padded | Rejection | corrected_id | corrected_passport | status_text | Status |
|---|---|---|---|---|---|---|---|---|
| AZ-01 | `"000000000"` | `"000000000"` | — | All-zeros check before pad | `""` | `""` | `"ת.ז. לא תקינה"` | IMPLEMENTED |
| AZ-02 | `"0"` | `"0"` | `"000000000"` | All-zeros padded | `""` | `""` | `"ת.ז. לא תקינה"` | IMPLEMENTED |
| AZ-03 | `"111111111"` | `"111111111"` | `"111111111"` | `len(set(padded))==1` | `""` | `""` | `"ת.ז. לא תקינה"` | IMPLEMENTED |
| AZ-04 | `"999999999"` | `"999999999"` | `"999999999"` | All-identical | `""` | `""` | `"ת.ז. לא תקינה"` | IMPLEMENTED |
| AZ-05 | `"123456789"` (all different) | `"123456789"` | `"123456789"` | Not rejected | checksum validated | depends | IMPLEMENTED |

#### 3.10.5 Digit Count — Too Short / Too Long

| Case ID | id_str | digit_count | Move? | corrected_id | corrected_passport | status_text | Status |
|---|---|---|---|---|---|---|---|
| DC-01 | `"123"` | 3 | Yes (< 4) | `""` | `"123"` | `"ת.ז. לא תקינה + הועברה לדרכון"` | IMPLEMENTED |
| DC-02 | `"1234"` | 4 | No | padded to `"000001234"` | depends | depends | IMPLEMENTED |
| DC-03 | `"123456789"` | 9 | No | `"123456789"` | depends | depends | IMPLEMENTED |
| DC-04 | `"1234567890"` | 10 | Yes (> 9) | `""` | `"1234567890"` | `"ת.ז. לא תקינה + הועברה לדרכון"` | IMPLEMENTED |
| DC-05 | `"12"` | 2 | Yes (< 4) | `""` | `"12"` | `"ת.ז. לא תקינה + הועברה לדרכון"` | IMPLEMENTED |
| DC-06 | `"1"` | 1 | Yes (< 4) | `""` | `"1"` | `"ת.ז. לא תקינה + הועברה לדרכון"` | IMPLEMENTED |

#### 3.10.6 Checksum Validation and Output

| Case ID | id_str | digits | padded | checksum valid | corrected_id | corrected_passport | status_text | Status |
|---|---|---|---|---|---|---|---|---|
| CS-01 | `"123456782"` | `"123456782"` | `"123456782"` | True (sum=50, 50%10=0) | `"123456782"` (original) | `""` | `"ת.ז. תקינה"` | IMPLEMENTED |
| CS-02 | `"123456782"` | same | same | True | `"123456782"` | `"AB123"` | `"ת.ז. תקינה + דרכון הוזן"` | IMPLEMENTED |
| CS-03 | `"123456789"` | `"123456789"` | `"123456789"` | False | `"123456789"` (padded) | `""` | `"ת.ז. לא תקינה"` | IMPLEMENTED |
| CS-04 | `"123456789"` | same | same | False | `"123456789"` | `"AB123"` | `"ת.ז. לא תקינה + דרכון הוזן"` | IMPLEMENTED |
| CS-05 | `"12345678"` (8 digits) | `"12345678"` | `"012345678"` | depends on padded checksum | padded `"012345678"` if invalid | `""` | depends | IMPLEMENTED |
| CS-06 | `"1234"` (4 digits) | `"1234"` | `"000001234"` | depends | padded if invalid | `""` | depends | IMPLEMENTED |

**Key rule:** When checksum is valid, `corrected_id = id_str` (original string). When invalid (but not moved), `corrected_id = padded` (9-digit zero-padded string).

#### 3.10.7 Passport Cleaning

**Code:** `IdentifierEngine.clean_passport()`

| Case ID | Input | Kept | Dropped | Output | Status |
|---|---|---|---|---|---|
| PC-01 | `"AB123"` | A,B,1,2,3 | — | `"AB123"` | IMPLEMENTED |
| PC-02 | `"AB 123"` | A,B,1,2,3 | space | `"AB123"` | IMPLEMENTED |
| PC-03 | `"AB-123"` | A,B,-,1,2,3 | — | `"AB-123"` | IMPLEMENTED |
| PC-04 | `"AB–123"` (en-dash) | A,B,–,1,2,3 | — | `"AB–123"` | IMPLEMENTED |
| PC-05 | `"AB!123"` | A,B,1,2,3 | `!` | `"AB123"` | IMPLEMENTED |
| PC-06 | `"יוסי123"` | י,ו,ס,י,1,2,3 | — | `"יוסי123"` | IMPLEMENTED |
| PC-07 | `""` | — | — | `""` | IMPLEMENTED |
| PC-08 | `None` | — | — | `""` | IMPLEMENTED |
| PC-09 | `"AB.123"` | A,B,1,2,3 | `.` | `"AB123"` | IMPLEMENTED |
| PC-10 | `"AB(123)"` | A,B,1,2,3 | `(`,`)` | `"AB123"` | IMPLEMENTED |

#### 3.10.8 Pipeline-Level Behavior

**Code:** `NormalizationPipeline.apply_identifier_normalization()`

| Case ID | id_number in row | passport in row | id_value | passport_value | Behavior | Status |
|---|---|---|---|---|---|---|
| PL-01 | No | No | — | — | Returns immediately; no corrected fields written | IMPLEMENTED |
| PL-02 | Yes | No | `None` | — | Both None/empty → `id_number_corrected = None`; no `identifier_status` written | IMPLEMENTED |
| PL-03 | Yes | Yes | `None` | `None` | Both None/empty → `id_number_corrected = None`, `passport_corrected = None`; no status | IMPLEMENTED |
| PL-04 | Yes | Yes | `"123456782"` | `"AB123"` | Engine called; results written; `identifier_status` written | IMPLEMENTED |
| PL-05 | Yes | No | `"123456782"` | — | `passport_value = None`; engine called with `(id, None)` | IMPLEMENTED |

**Note on PL-02/PL-03:** When both values are None/empty, the pipeline returns early. `identifier_status` is NOT written to the row. The status key will be absent.

---

### 3.11 MosadID Extraction

**Code:** `mosad_id_scanner.scan_mosad_id()` — `mosad_id_scanner.py`

| Case ID | Condition | Behavior | Output | Status |
|---|---|---|---|---|
| MI-01 | Cell contains `"מספר מוסד"`, right neighbour has `"12345"` | Label matched; right checked first | `"12345"` | IMPLEMENTED |
| MI-02 | Cell contains `"מספר מוסד"`, right neighbour empty, left has `"12345"` | Right empty; left checked | `"12345"` | IMPLEMENTED |
| MI-03 | Cell contains `"מספר מוסד"`, both neighbours empty | No non-empty neighbour | `None` | IMPLEMENTED |
| MI-04 | Cell contains `"institution id"` | English label matched | value from neighbour | IMPLEMENTED |
| MI-05 | Cell contains `"mosadid"` | Matched | value from neighbour | IMPLEMENTED |
| MI-06 | No label cell anywhere in sheet | Scan completes without match | `None` | IMPLEMENTED |
| MI-07 | Multiple label cells in sheet | First match (top-to-bottom, left-to-right) wins | First found value | IMPLEMENTED |
| MI-08 | Label cell at column 1 (no left neighbour) | `neighbour_col < 1` → skipped | Right neighbour checked only | IMPLEMENTED |
| MI-09 | Label cell at last column (no right neighbour) | `neighbour_col > max_col` → skipped | Left neighbour checked only | IMPLEMENTED |
| MI-10 | MosadID value is `None` | `_coerce_value(None)` → `None` | Not returned | IMPLEMENTED |
| MI-11 | MosadID value is `""` | `_coerce_value("")` → `None` | Not returned | IMPLEMENTED |
| MI-12 | MosadID found; injected into rows | `apply_derived_columns` | All rows get `MosadID` key | IMPLEMENTED |
| MI-13 | MosadID not found; no rows have it | `mosad_id_has_value = False` | `MosadID` column not shown | IMPLEMENTED |

---

### 3.12 Session Edit / Delete / Re-normalize Behavior

**Code:** `EditService`, `NormalizationService`, `SessionService`

#### 3.12.1 Cell Edit

| Case ID | Condition | Behavior | Side Effects | Status |
|---|---|---|---|---|
| SE-01 | Valid `row_index`, valid `field_name` | `sheet.rows[row_index][field_name] = new_value` | `record.edits[(sheet, row, field)] = new_value` | IMPLEMENTED |
| SE-02 | `row_index < 0` | HTTP 400 | No change | IMPLEMENTED |
| SE-03 | `row_index >= len(rows)` | HTTP 400 | No change | IMPLEMENTED |
| SE-04 | `field_name` not in row dict | HTTP 400 | No change | IMPLEMENTED |
| SE-05 | Edit a `_corrected` field | Allowed; no re-normalization | `_corrected` value updated in memory | IMPLEMENTED |
| SE-06 | Edit an original field | Allowed; no re-normalization | Original updated; `_corrected` retains old value | IMPLEMENTED |
| SE-07 | `workbook_dataset is None` | HTTP 500 | No change | IMPLEMENTED |
| SE-08 | `_normalization_failures` key in row | Stripped from `updated_row` response | Not visible in response | IMPLEMENTED |

#### 3.12.2 Row Deletion

| Case ID | Condition | Behavior | Side Effects | Status |
|---|---|---|---|---|
| SD-01 | Valid single index | `sheet.rows.pop(idx)` | Row gone from memory; disk unchanged | IMPLEMENTED |
| SD-02 | Multiple valid indices | Deduplicated, sorted, removed in reverse order | All removed atomically | IMPLEMENTED |
| SD-03 | Duplicate indices in request | Deduplicated before validation | Each row deleted once | IMPLEMENTED |
| SD-04 | Any index out of range | HTTP 400; NO rows deleted | All-or-nothing | IMPLEMENTED |
| SD-05 | Empty `row_indices` list | HTTP 400 | No change | IMPLEMENTED |
| SD-06 | Delete all rows | All rows removed | Sheet has 0 rows | IMPLEMENTED |
| SD-07 | Delete then re-normalize | Re-normalize re-extracts from disk | Deleted rows reappear | IMPLEMENTED |

#### 3.12.3 Re-normalize Behavior

| Case ID | Condition | Behavior | Status |
|---|---|---|---|
| RN-01 | Normalize called with no `?sheet=` param | All sheets re-extracted from disk; all normalized | IMPLEMENTED |
| RN-02 | Normalize called with `?sheet=SheetName` | Only that sheet re-extracted and normalized; others unchanged | IMPLEMENTED |
| RN-03 | Manual edits exist before normalize | Re-extraction from disk discards all in-memory edits | IMPLEMENTED |
| RN-04 | `record.edits` dict | Populated on edit; never read back; never replayed | IMPLEMENTED |
| RN-05 | Normalize called twice | Second call re-extracts fresh; first normalization results discarded | IMPLEMENTED |
| RN-06 | `workbook_dataset is None` when normalize called | Auto-loads all sheets from disk first | IMPLEMENTED |
| RN-07 | Sheet not found on disk during normalize | HTTP 404 | IMPLEMENTED |

---

### 3.13 Export Behavior

**Code:** `ExportService.export()`, `visible_rows()`, `canonical_sheet_name()`, `headers_for_sheet()`, `EXPORT_MAPPING` — `export_service.py`

#### 3.13.1 Sheet Name Canonicalization

| Case ID | Source Sheet Name | Canonical Name | Schema Used | Status |
|---|---|---|---|---|
| EX-01 | `"דיירים יחידים"` | `DayarimYahidim` | 14-column (no Dira) | IMPLEMENTED |
| EX-02 | `"דיירים"` | `DayarimYahidim` | 14-column | IMPLEMENTED |
| EX-03 | `"מתגוררים במשקי בית"` | `MeshkeyBayt` | 15-column (with Dira) | IMPLEMENTED |
| EX-04 | `"משקי בית"` | `MeshkeyBayt` | 15-column | IMPLEMENTED |
| EX-05 | `"מתגוררים"` | `MeshkeyBayt` | 15-column | IMPLEMENTED |
| EX-06 | `"אנשי צוות ובני משפחותיהם"` | `AnasheyTzevet` | 15-column | IMPLEMENTED |
| EX-07 | `"אנשי צוות"` | `AnasheyTzevet` | 15-column | IMPLEMENTED |
| EX-08 | `"צוות"` | `AnasheyTzevet` | 15-column | IMPLEMENTED |
| EX-09 | `"Sheet1"` (no match) | `"Sheet1"` (unchanged) | 14-column (default) | IMPLEMENTED |
| EX-10 | `"  דיירים  "` (extra spaces) | `DayarimYahidim` (NFC + whitespace collapse) | 14-column | IMPLEMENTED |

#### 3.13.2 Field Mapping — Corrected-Only Rule

| Case ID | Condition | Export Cell Value | Status |
|---|---|---|---|
| FM-01 | `first_name_corrected` present and non-empty | Value of `first_name_corrected` | IMPLEMENTED |
| FM-02 | `first_name_corrected` absent (no normalization) | `None` → blank cell | IMPLEMENTED |
| FM-03 | `first_name_corrected` is `""` | `None` → blank cell | IMPLEMENTED |
| FM-04 | `first_name` present but `first_name_corrected` absent | No fallback; blank cell | IMPLEMENTED |
| FM-05 | `MosadID` present in row | Value written | IMPLEMENTED |
| FM-06 | `SugMosad` present in row | Value written (but never populated by pipeline) | IMPLEMENTED |
| FM-07 | `SugMosad` absent | Blank cell | IMPLEMENTED |
| FM-08 | `MisparDiraBeMosad` present in row | Value written | IMPLEMENTED |
| FM-09 | `MisparDiraBeMosad` absent | Blank cell | IMPLEMENTED |
| FM-10 | `gender_corrected` = `1` (int) | `1` written as integer | IMPLEMENTED |
| FM-11 | `gender_corrected` = `2` (int) | `2` written as integer | IMPLEMENTED |
| FM-12 | `birth_year_corrected` = `1980` (int) | `1980` written | IMPLEMENTED |

#### 3.13.3 Row Filtering in Export

Same logic as UI (`visible_rows()` mirrors `WorkbookService.get_sheet_data()`):

| Case ID | Condition | Included in Export | Status |
|---|---|---|---|
| EF-01 | Row with all original cols empty | No | IMPLEMENTED |
| EF-02 | Row with at least one non-empty original col | Yes | IMPLEMENTED |
| EF-03 | First row is all-numeric (helper row) | No | IMPLEMENTED |
| EF-04 | Second row is all-numeric (not first) | Yes | IMPLEMENTED |
| EF-05 | Row deleted via UI before export | No (gone from memory) | IMPLEMENTED |
| EF-06 | Row deleted, then re-normalized before export | Yes (reappears from disk) | IMPLEMENTED |

#### 3.13.4 Export Without Prior Normalization

| Case ID | Condition | Behavior | Status |
|---|---|---|---|
| EN-01 | Export called, `workbook_dataset is None` | Auto-extracts from disk; no `_corrected` fields | IMPLEMENTED |
| EN-02 | Export called after extraction but before normalize | `_corrected` fields absent → all personal data columns blank | IMPLEMENTED |
| EN-03 | Export called after normalize | `_corrected` fields present → data written | IMPLEMENTED |

#### 3.13.5 Output File

| Case ID | Condition | Behavior | Status |
|---|---|---|---|
| OF-01 | Normal export | `{stem}_normalized_{YYYYMMDD_HHMMSS}.xlsx` in `output/` | IMPLEMENTED |
| OF-02 | Source was `.xlsm` | Output is always `.xlsx` | IMPLEMENTED |
| OF-03 | Two exports in same second | Different timestamps → no collision | IMPLEMENTED |
| OF-04 | Sheet direction | `ws.sheet_view.rightToLeft = True` | IMPLEMENTED |
| OF-05 | Header alignment | `Alignment(horizontal="right")` on header cells | IMPLEMENTED |
| OF-06 | No pink highlights in export | `ExportService` does not apply any cell formatting | IMPLEMENTED |

---

---

## 4. Path Differences: Web/JSON vs Direct-Excel

| Area | Web / JSON Path | Direct-Excel / CLI Path |
|---|---|---|
| **File modification** | Original never touched; new `.xlsx` created on export | Workbook modified in-place; corrected columns inserted |
| **Date format detection** | Always `DateFormatPattern.DDMM` (hardcoded in `_normalize_date_field`) | `detect_date_format_pattern()` scans data; can return MMDD |
| **Entry-before-birth check** | NOT performed | Performed by `_validate_entry_vs_birth()` after both date groups written |
| **Date header search terms** | `ExcelReader` uses `FIELD_KEYWORDS["birth_date"]` = `["תאריך לידה", "birth date", ...]` | `DateFieldProcessor` searches `["תאריך לידה"]` (birth) and `["תאריך כניסה למוסד"]` (entry) — more restrictive |
| **Multiple date groups per sheet** | JSON path uses first group only (from `detect_date_groups`) | Direct-Excel path processes ALL groups found by `_collect_date_groups` |
| **Pink cell highlighting** | Not applied | Applied by `ExcelWriter.highlight_changed_cells()` |
| **Status cell formatting** | Not applied | Age warning → yellow; other errors → pink+bold |
| **Column insertion** | Not done; corrected values go into new JSON keys | `insert_cols()` called in-place; column positions shift |
| **Pattern detection sample** | First 10 rows with both fields (from `normalize_dataset`) | First 5 rows (from `NameFieldProcessor.detect_father_name_pattern`) |
| **Pattern detection input** | Raw values from rows (then `normalize_name` called on each) | Already-normalized values |
| **`_corrected` field format (dates)** | Split: integer components; Single: `"DD/MM/YYYY"` string | Integer components written to cells; number format `"0"` applied |
| **Export schema** | Fixed 14/15-column schema; corrected-only | VBA-parity: reads `"- מתוקן"` headers; exports corrected values |
| **MosadID in export** | From `scan_mosad_id()` metadata | From tracking dict only (not scanned) |
| **Session state** | In-memory; lost on restart | No session; file modified directly |
| **`.xlsm` handling** | `data_only=True`; VBA not executed | `keep_vba=True`; VBA preserved in saved file |
| **`SugMosad`** | Never populated | Never populated |
| **`MisparDiraBeMosad`** | Not extracted from source | Not extracted from source |

---

## 5. Observed Risky Current Behaviors

1. **Manual edits are silently discarded on re-normalize.** `NormalizationService.normalize()` re-extracts from disk, replacing the in-memory dataset. `record.edits` is populated but never replayed. A user who edits cells and then re-normalizes loses all edits with no warning.

2. **`gender_corrected` is an `int` (1 or 2), not a string.** The export writes it as an integer. The UI displays it as `"1"` or `"2"`. If downstream systems expect `"ז"` or `"נ"`, they will receive a number.

3. **`"נ"` substring match in gender.** Any string containing the Hebrew letter Nun (e.g., `"נ/א"`, `"זכר/נקבה"`, `"ינואר"`) will be classified as female. The check is `pattern.lower() in value_str`, not a whole-word match.

4. **Float ID values.** An Excel cell containing `123456789.0` (float) will produce `id_str = "123456789.0"`. The `.` is not a digit or dash, so the ID is moved to passport. This is likely unintended for numeric Excel cells.

5. **Stage B fires on partial-word matches.** If the last name is `"כהן"` and the first name is `"כהנים יוסי"`, Stage A does not match (word-boundary aware), so Stage B fires and removes the first token (`"כהנים"`), leaving `"יוסי"`. The removed token was not the last name.

6. **Pattern detection uses first 10 rows only.** If the first 10 rows are atypical (e.g., all have the last name embedded, but the rest don't), the pattern is applied to all rows including those where it is wrong.

7. **`birth_date_status` / `identifier_status` absent when input is None/empty.** When `date_val is None or == ""` (single date path), the pipeline returns early without writing `birth_date_status`. When both id and passport are None/empty, `identifier_status` is not written. The status key is absent from the row dict, which differs from the case where the engine runs and returns an empty status string.

8. **Numeric helper-row filter only checks the first row.** If the second row is also all-numeric, it is not filtered. Only the first data row is checked.

9. **`_serial` column header is an internal key name.** The synthetic serial column appears in the UI with the header `_serial`, which is not user-friendly.

10. **Export without normalization produces a mostly-blank file.** There is no warning or guard. The user receives a valid `.xlsx` with correct structure but blank personal data columns.

11. **`record.edits` dict grows unboundedly.** Every cell edit appends to `record.edits`. There is no size limit or cleanup. For large datasets with many edits, this dict can grow large, but since it is never read back, it is pure memory waste.

12. **Two-digit year expansion is date-of-run dependent.** The cutoff year changes each year. A year value of `"26"` means 2026 in 2026 but will mean 1926 in 2027. Archived data processed in different years will produce different results.

13. **`validate_entry_before_birth` in `DateEngine` is dead code in the web path.** The method exists and is correct, but `NormalizationPipeline` never calls it. The cross-validation only runs in the direct-Excel CLI path.

14. **`SugMosad` is always blank in all export paths.** The column exists in the schema but no code populates it. Any downstream system expecting this field will always receive an empty value.

---

*Generated from codebase analysis. All cases derived from implemented code only.*
