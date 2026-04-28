# Excel standardization System — Current-State Specification

> **Document type:** Current-state analysis and behavioral specification.
> Based entirely on the real codebase as it exists today.
> Does not describe a future ideal system.
> Status markers: **[IMPLEMENTED]**, **[PARTIAL]**, **[NOT IMPLEMENTED]**, **[AMBIGUOUS]**

---

## Table of Contents

1. [General System Overview](#1-general-system-overview)
2. [Architecture Overview](#2-architecture-overview)
3. [Implemented vs Partial vs Missing — by Area](#3-implemented-vs-partial-vs-missing)
4. [Sheet Loading — Detailed Behavior](#4-sheet-loading--detailed-behavior)
5. [Name and Text standardization — Detailed Rules](#5-name-and-text-standardization--detailed-rules)
6. [Last-Name Removal from First Name and Father Name](#6-last-name-removal-from-first-name-and-father-name)
7. [Date standardization — Detailed Rules](#7-date-standardization--detailed-rules)
8. [Identifier standardization — Detailed Rules](#8-identifier-standardization--detailed-rules)
9. [UI Behavior](#9-ui-behavior)
10. [Export Behavior](#10-export-behavior)
11. [Installer and Launcher Behavior](#11-installer-and-launcher-behavior)
12. [Edge-Case Catalogue](#12-edge-case-catalogue)
13. [Gaps, Limitations, and Open Questions](#13-gaps-limitations-and-open-questions)

---

## 1. General System Overview

### What the system does

The system takes Excel workbooks (`.xlsx` / `.xlsm`) containing person records — residents, staff, and family members of Israeli institutions — and normalizes inconsistent data into a standardized format. It is a Python reimplementation of a legacy VBA macro system, designed to produce identical output.

The four standardization domains are:

| Domain | What it does |
|---|---|
| **Names** | Cleans first name, last name, father's name: removes diacritics, wrong-language characters, honorific titles, and optionally strips the last name from the first/father name field |
| **Gender** | Maps any representation (Hebrew text, English text, numeric) to `1` (male) or `2` (female) |
| **Dates** | Parses birth date and entry date from split columns (year/month/day) or single string cells; validates against business rules; writes Hebrew status messages |
| **Identifiers** | Validates Israeli ID numbers (Luhn-style checksum), cleans passport values, moves non-ID values to the passport field, writes Hebrew status messages |

### Main user flows

**Web app flow (primary):**
1. User opens the app in a browser (launched via `launcher.py` or `uvicorn`)
2. Uploads one or more `.xlsx` / `.xlsm` files
3. Browses sheet data in a grid view
4. Clicks "Run standardization" — all sheets are processed
5. Reviews corrected values inline; edits cells if needed; deletes rows if needed
6. Clicks "Export / Download" — receives a normalized `.xlsx` file

**CLI flow (secondary, developer-facing):**
- `python -m excel_standardization.cli path/to/file.xlsx` — runs the legacy direct-Excel processor path, modifying the workbook in-place

### Main technical layers

```
Browser (vanilla JS SPA)
        ↕ HTTP/JSON
FastAPI web app  (webapp/)
        ↕
Session registry (in-memory dict, process lifetime)
        ↕
standardizationService / WorkbookService / ExportService / EditService
        ↕
ExcelToJsonExtractor  →  standardizationPipeline  →  ExportService
        ↕                        ↕
    ExcelReader            NameEngine / GenderEngine
                           DateEngine / IdentifierEngine
                           TextProcessor
```

The system has two parallel processing paths:

- **JSON pipeline** (used by the web app): Excel → JSON (`SheetDataset`) → normalize in memory → export to new `.xlsx`
- **Direct-Excel processor path** (legacy, used by CLI): Excel → insert corrected columns in-place → save

The web app exclusively uses the JSON pipeline. The CLI uses the direct-Excel path. Both paths share the same engine classes.

---

## 2. Architecture Overview

### Upload / Session / Workbook handling

**Upload** (`UploadService`):
- Accepts `.xlsx` and `.xlsm` only; rejects everything else with HTTP 400
- Generates a UUID session ID
- Saves the original file to `uploads/{session_id}.ext` (never modified)
- Copies it to `work/{session_id}.ext` (the working copy)
- Opens the workbook read-only to get sheet names; does NOT extract data yet
- Creates a `SessionRecord` in the in-memory registry with `workbook_dataset=None`

**Session registry** (`SessionService`):
- Module-level Python dict — shared across all requests in the process
- No persistence: all sessions are lost on server restart
- No expiry or cleanup mechanism
- No locking (relies on single-threaded Uvicorn)
- Status field: `"uploaded"` → `"normalized"` (only two states)

**Workbook dataset** (`WorkbookDataset` / `SheetDataset`):
- Lazy loading: a sheet is only extracted from disk when first requested via `GET /api/workbook/{id}/sheet/{name}`
- After extraction, the `SheetDataset` is stored in `record.workbook_dataset.sheets`
- standardization re-extracts from disk (fresh read), then replaces the in-memory sheet
- Manual edits are applied directly to the in-memory `SheetDataset.rows`
- Edits are recorded in `record.edits` dict but this dict is never used for anything (not replayed on export)

### standardization layer

**`ExcelToJsonExtractor`**:
- Opens workbook with `data_only=True` (formula results, not formula strings)
- Calls `ExcelReader.detect_table_region()` to find where the data table starts
- Calls `ExcelReader.detect_columns()` to map header text to field names
- Extracts each data row into a `JsonRow` dict
- Formula cells that return `#ERROR` → stored as `None`
- Formula cells that return the formula string (unevaluated, starts with `=`) → stored as `None`
- Merged cells: openpyxl returns the top-left cell value automatically

**`standardizationPipeline`**:
- Operates on `SheetDataset` (list of `JsonRow` dicts)
- Non-destructive: original fields are never modified; corrected values go into `field_corrected` keys
- Processing order per row: names → gender → dates → identifiers
- Pattern detection (last-name removal) runs once per dataset on the first 10 rows, then is cached on the pipeline instance for all rows
- Failed standardizations are recorded in `_standardization_failures` key on the row (stripped before display)

**`ExcelReader` — table detection**:
- Scans up to 30 rows to find the header row by scoring each row (keyword matches, text density)
- Detects 1-row or 2-row headers (2-row = parent header + year/month/day sub-headers)
- Detects and skips a column-index helper row (row of sequential integers immediately after headers)
- Columns not matching any known keyword are passed through with a sanitised version of their header text as the field name
- Columns whose header contains "מתוקן" or "corrected" are ignored (already-processed columns)

### Export layer

**`ExportService`** (web app export):
- Creates a new workbook; one sheet per source sheet
- Sheet names are mapped to canonical names: `DayarimYahidim`, `MeshkeyBayt`, `AnasheyTzevet`
- Fixed column schema per sheet type (14 or 15 columns)
- Writes only `*_corrected` field values — no fallback to originals
- Applies the same row/column filters as the UI (empty rows dropped, numeric helper row dropped)
- Sets `rightToLeft = True` on each sheet
- Output filename: `{original_stem}_normalized_{timestamp}.xlsx`
- Saved to `output/` directory; served as a file download

**`ExportEngine`** (legacy VBA-parity engine, used by CLI):
- Two modes: `export_from_augmented_workbook` (after direct-Excel processors) and `export_from_normalized_dataset` (after JSON pipeline)
- Detects corrected columns by scanning for headers containing `"- מתוקן"`
- Row validity: a row is exported if ANY of `ShemPrati`, `ShemMishpaha`, `ShemHaAv`, `MisparZehut`, `Darkon` is non-empty
- `MosadID`, `SugMosad`, `MisparDiraBeMosad` are only written if present in the tracking dictionary

### Web UI

Single-page application in vanilla JavaScript (`webapp/static/app.js`). No framework, no build step, fully offline-capable.

Key UI components:
- File upload form (supports multiple files)
- Session switcher (tab per uploaded file, shows `✓` badge after standardization)
- Sheet selector (tab per sheet)
- Data grid (HTML table with inline editing)
- Action bar: "Run standardization", "Export / Download", "Delete rows"

### Installer / Launcher

- `launcher.py`: PyInstaller entry point; starts Uvicorn on a free port (preferred: 8765), opens Chrome in app mode (`--app=URL --new-window`), falls back to system default browser
- `installer/Excelstandardization.iss`: Inno Setup script; produces `Excelstandardization_Setup_1.0.1.exe`
- Runtime data directories: `%LOCALAPPDATA%\Excelstandardization\{uploads,work,output}`
- Log file: `%LOCALAPPDATA%\Excelstandardization\app.log`
- Windows 10+ x64 only; requires admin for installation

---

## 3. Implemented vs Partial vs Missing

### Upload and session management

| Feature | Status | Notes |
|---|---|---|
| `.xlsx` / `.xlsm` upload | **IMPLEMENTED** | Extension validated; workbook opened to verify |
| Session creation with UUID | **IMPLEMENTED** | |
| Original file preserved (never modified) | **IMPLEMENTED** | Separate `uploads/` copy |
| Lazy sheet extraction | **IMPLEMENTED** | Sheets loaded on first access |
| Session persistence across restarts | **NOT IMPLEMENTED** | In-memory only; restart = data loss |
| Session expiry / cleanup | **NOT IMPLEMENTED** | Sessions accumulate for process lifetime |
| Multi-user / concurrent access | **NOT IMPLEMENTED** | No locking; single-threaded assumption |
| Edit replay on export | **NOT IMPLEMENTED** | `record.edits` dict is populated but never read back |

### standardization

| Feature | Status | Notes |
|---|---|---|
| Name standardization (clean) | **IMPLEMENTED** | Full pipeline: diacritics, language detection, char filtering, token removal |
| Last-name removal from first name | **IMPLEMENTED** | Two-stage: substring + positional fallback |
| Last-name removal from father name | **IMPLEMENTED** | Same two-stage logic |
| Gender standardization | **IMPLEMENTED** | Maps to 1/2; Hebrew, English, numeric inputs |
| Date parsing — split columns | **IMPLEMENTED** | Year/month/day columns |
| Date parsing — single string | **IMPLEMENTED** | ISO, DD/MM/YYYY, DDMMYYYY, month names |
| Date business rules | **IMPLEMENTED** | Future date, pre-1900, age > 100 |
| Entry-before-birth cross-validation | **IMPLEMENTED** | In direct-Excel path; **PARTIAL** in JSON pipeline (status written but no cross-sheet check) |
| Israeli ID checksum validation | **IMPLEMENTED** | Luhn-style algorithm |
| Passport cleaning | **IMPLEMENTED** | Keeps digits, ASCII letters, Hebrew letters, dashes |
| ID → passport migration | **IMPLEMENTED** | Non-digit IDs, too-short/too-long IDs moved to passport |
| Date format detection (DDMM vs MMDD) | **IMPLEMENTED** | Heuristic based on values > 12 |
| MosadID scanning | **IMPLEMENTED** | Scans for label/value pair outside main table |
| Serial number injection | **IMPLEMENTED** | Synthetic `_serial` if no source column; auto-fill blanks |
| `SugMosad` field | **NOT IMPLEMENTED** | Column exists in export schema; always blank in web export |

### UI

| Feature | Status | Notes |
|---|---|---|
| Multi-file upload | **IMPLEMENTED** | Sequential uploads, one session per file |
| File switching (session switcher) | **IMPLEMENTED** | Restores last-viewed sheet per session |
| Sheet tabs | **IMPLEMENTED** | |
| Data grid with original + corrected columns | **IMPLEMENTED** | |
| Inline cell editing | **IMPLEMENTED** | Click any non-status cell |
| Single-row delete | **IMPLEMENTED** | ✕ button per row |
| Multi-row delete (checkbox) | **IMPLEMENTED** | Select-all + bulk delete |
| standardization badge (`✓`) on file tab | **IMPLEMENTED** | |
| Bulk ZIP export (multi-file) | **PARTIAL** | Code exists but UI buttons are commented out |
| Undo / redo | **NOT IMPLEMENTED** | |
| Confirmation dialog before delete | **NOT IMPLEMENTED** | Deletes immediately |
| Pagination for large sheets | **NOT IMPLEMENTED** | All rows rendered at once |
| Column sorting / filtering | **NOT IMPLEMENTED** | |
| Status column not editable | **IMPLEMENTED** | `_status` columns skip click handler |

### Export

| Feature | Status | Notes |
|---|---|---|
| Fixed-schema export (14/15 columns) | **IMPLEMENTED** | |
| Corrected-only values | **IMPLEMENTED** | No fallback to originals |
| RTL sheet direction | **IMPLEMENTED** | |
| Sheet name canonicalization | **IMPLEMENTED** | Hebrew → English canonical names |
| Row filtering (empty rows, helper row) | **IMPLEMENTED** | |
| `MosadID` in export | **IMPLEMENTED** | From sheet metadata scan |
| `SugMosad` in export | **NOT IMPLEMENTED** | Always blank |
| `MisparDiraBeMosad` in export | **PARTIAL** | In schema for MeshkeyBayt/AnasheyTzevet; only populated if present in source rows |
| Bulk ZIP export | **PARTIAL** | API endpoint exists; UI disabled |
| Export without prior standardization | **IMPLEMENTED** | Auto-extracts from disk; writes uncorrected values (corrected fields absent → blank) |

---

## 4. Sheet Loading — Detailed Behavior

### Header detection

`ExcelReader.detect_table_region()` scans up to 30 rows and scores each row:

- +2 if ≥ 3 non-empty cells
- +1 if ≥ 5 non-empty cells
- +2 if ≥ 70% of non-empty cells are text (not pure numbers)
- +2 per cell that contains a known field keyword

The row with the highest score (minimum 3) is selected as the header row. If no row scores ≥ 3, the sheet is skipped.

**Two-row header detection:** After selecting the best row, the system checks whether:
- The selected row is itself a sub-header row (contains שנה/חודש/יום/year/month/day keywords), in which case the row above becomes the parent header
- The row below the selected row looks like a sub-header row (scores ≥ 2 on sub-header scoring)
- The row above has merged cells spanning multiple columns with date-related keywords

When two header rows are detected, `header_rows = 2` and `data_start_row = header_start_row + 2`.

### Column-index helper row

Immediately after the header rows, some Excel forms include a row of sequential integers (1, 2, 3…) labeling column positions. This row is detected and skipped if:
- Every non-null cell is a positive integer
- All values are ≤ the table's end column
- At least 3 values are present
- Values are distinct and consecutive (no gap > 1)

This detection runs at extraction time and also at display time in `WorkbookService.get_sheet_data()`.

### Empty row filtering

Two separate empty-row filters are applied:

1. **Extraction time** (`skip_empty_rows=False` by default): the extractor does NOT skip empty rows during extraction. All rows including blanks are stored in `SheetDataset.rows`.

2. **Display/export time**: `WorkbookService.get_sheet_data()` and `ExportService.visible_rows()` both drop rows where every original-column cell is `None`, empty string, or whitespace-only. The check is against original source columns only — `_corrected` and `_status` columns are excluded from this check.

### Visible columns and column ordering

`WorkbookService.get_sheet_data()` builds `display_columns` in this order:

1. For each original field (in Excel left-to-right order):
   a. The original field itself
   b. Its `_corrected` column (if present in any row)
   c. Its status column, if this field is the **rightmost** member of a status group in the sheet

2. Any remaining keys not yet placed (unexpected extras)

3. Serial number column prepended at position 0
4. MosadID column at position 1 (if any row has a non-empty MosadID)

**Status groups:**
- `identifier_status` anchors to the rightmost of `{id_number, passport}` in the sheet
- `birth_date_status` anchors to the rightmost of `{birth_year, birth_month, birth_day, birth_date}`
- `entry_date_status` anchors to the rightmost of `{entry_year, entry_month, entry_day, entry_date}`

### File switching

When the user switches between uploaded files (sessions):
- The outgoing session's `lastSheet` is saved to the session record
- The incoming session's `lastSheet` is restored and that sheet is loaded
- The grid is cleared and re-rendered with the new session's data
- `state.selectedRows` is cleared

### Row deletion behavior

- Deletion is applied to the in-memory `SheetDataset.rows` list only
- The working copy file on disk is NOT modified
- Deletion is permanent within the session (no undo)
- All indices are validated before any deletion occurs (all-or-nothing)
- The frontend adjusts `selectedRows` indices after deletion to account for the shift

### Multi-file behavior

- Each uploaded file gets its own independent session
- Sessions share no state
- The session switcher shows one tab per file
- standardization, editing, and export are all per-session
- Bulk ZIP export is implemented in the API but the UI buttons are commented out

---

## 5. Name and Text standardization — Detailed Rules

All name standardization goes through `TextProcessor.clean_name()`. The pipeline is strictly ordered and cannot be reordered.

### Order of operations

```
1. SafeToString + strip zero-width characters
2. Diacritic removal
3. Arabic-Indic digit standardization (٠١٢٣٤٥٦٧٨٩ → 0-9)
4. Language detection (Hebrew vs English vs Mixed)
5. Character filtering
6. Space standardization (trim + collapse)
7. Unwanted token removal
```

### Step 1 — SafeToString + zero-width stripping

- `None` → `""` (returns immediately)
- Any value is converted via `str()`
- Zero-width characters stripped: U+200B, U+200C, U+200D, U+200E, U+200F, U+202A–U+202E, U+FEFF

### Step 2 — Diacritic removal

Explicit character map. Covers common Latin diacritics (à→a, é→e, ü→u, ñ→n, ç→c, etc.) and Cyrillic ё→e. Hebrew characters are not affected.

### Step 3 — Arabic-Indic digits

`٠١٢٣٤٥٦٧٨٩` → `0123456789` via `str.maketrans`. Applied before language detection so digits don't skew the count.

### Step 4 — Language detection

Counts Hebrew letters (U+05D0–U+05EA) vs ASCII letters (A–Z, a–z). Hebrew wins on a tie. Returns `HEBREW`, `ENGLISH`, or `MIXED` (when both counts are zero).

```
"יוסי כהן"     → HEBREW
"John Smith"   → ENGLISH
"יוסי Smith"   → HEBREW (Hebrew count ≥ English count)
"123"          → MIXED  (no letters at all)
""             → MIXED
```

### Step 5 — Character filtering

Rules applied per character:

| Character type | HEBREW mode | ENGLISH mode | MIXED mode |
|---|---|---|---|
| Space | kept | kept | kept |
| Hyphen-like chars (-, –, —, −, etc.) | → space | → space | → space |
| Hebrew letter (U+05D0–U+05EA) | kept | **dropped** | kept |
| ASCII letter (A–Z, a–z) | **dropped** | kept | kept |
| Digit | **dropped** | **dropped** | **dropped** |
| Any other character | **dropped** | **dropped** | **dropped** |

**Key consequences:**
- Digits are always dropped from names, regardless of language
- Symbols (punctuation, brackets, etc.) are always dropped
- Geresh (`'`, U+05F3) and gershayim (`"`, U+05F4) are dropped (they are not in the Hebrew letter range U+05D0–U+05EA)
- Regular ASCII quotes (`'`, `"`) are dropped
- Hyphens of all kinds become spaces
- In Hebrew mode, English letters are dropped; in English mode, Hebrew letters are dropped

### Step 6 — Space standardization

`" ".join(text.split())` — trims leading/trailing whitespace and collapses multiple internal spaces to one.

### Step 7 — Unwanted token removal

Applied after character filtering, so punctuation has already been removed.

**Hebrew tokens removed** (whole-word match, space-padded):

| Original form | After char filtering | Removed as |
|---|---|---|
| ז"ל | זל | `זל` |
| זצ"ל | זצל | `זצל` |
| זיע"א | זיעא | `זיעא` |
| הי"ד | היד | `היד` |
| שליט"א | שליטא | `שליטא` |
| ד"ר | דר | `דר` |
| רבי | רבי | `רבי` |
| ר (abbreviated rabbi) | ר | `ר` |
| ברד, ברמ, בראא, בראש, בימ, ברדא, ברי | same | same |

**English titles removed** (after char filtering removes the trailing dot):
`mr`, `mrs`, `ms`, `dr`, `prof`, `jr`, `sr`, `iii`, `iv`

### Examples

| Input | Language | Output |
|---|---|---|
| `"יוסי כהן"` | HEBREW | `"יוסי כהן"` |
| `"  יוסי   כהן  "` | HEBREW | `"יוסי כהן"` |
| `"יוסי-כהן"` | HEBREW | `"יוסי כהן"` (hyphen → space) |
| `"יוסי כהן ז\"ל"` | HEBREW | `"יוסי כהן"` (זל removed) |
| `"ד\"ר יוסי כהן"` | HEBREW | `"יוסי כהן"` (דר removed) |
| `"John Smith"` | ENGLISH | `"John Smith"` |
| `"Dr. John Smith"` | ENGLISH | `"John Smith"` (dr removed) |
| `"John Smith Jr."` | ENGLISH | `"John Smith"` (jr removed) |
| `"יוסי123"` | HEBREW | `"יוסי"` (digits dropped) |
| `"יוסי Smith"` | HEBREW | `"יוסי"` (English letters dropped in Hebrew mode) |
| `"יוסי'"` | HEBREW | `"יוסי"` (geresh dropped) |
| `"יוסי\u200b"` | HEBREW | `"יוסי"` (zero-width stripped) |
| `""` | MIXED | `""` |
| `None` | — | `""` |
| `"123"` | MIXED | `""` (digits dropped, nothing left) |
| `"à la"` | ENGLISH | `"a la"` (diacritic removed) |

### What is NOT normalized

- Capitalization of English names is preserved (no title-casing)
- Hebrew final letters (ך, ם, ן, ף, ץ) are kept as-is; `fix_hebrew_final_letters()` exists but is not called in the main pipeline
- No spell-checking or dictionary lookup
- No transliteration between Hebrew and English

---

## 6. Last-Name Removal from First Name and Father Name

### Overview

Some source workbooks store the last name inside the first name or father name field (e.g., `"כהן יוסי"` where `"כהן"` is the last name). The system detects this pattern and removes the last name from those fields.

This logic applies to both `first_name` and `father_name`. It does NOT apply to `last_name` itself (last name is only cleaned, never modified by removal).

### Pattern detection

Detection runs once per dataset (not per row), using the first 10 rows that have both the relevant field and `last_name` populated.

`detect_father_name_pattern()` and `detect_first_name_pattern()` use the same algorithm:

1. Sample up to 5 rows where both fields are non-empty
2. Count how many times the last name appears as a substring of the first/father name (`contain`)
3. Count how many times it appears as the **first token** (`first_pos`)
4. Count how many times it appears as the **last token** (`last_pos`)

Decision:
- If `contain < 3` → `FatherNamePattern.NONE` (no removal)
- If `first_pos >= 3` → `FatherNamePattern.REMOVE_FIRST`
- If `last_pos >= 3` → `FatherNamePattern.REMOVE_LAST`
- Otherwise → `FatherNamePattern.NONE`

**Important:** The sample uses the **cleaned** (normalized) versions of both fields, not the raw originals.

### Two-stage removal

Once a pattern is detected, each row goes through two-stage removal:

#### Stage A — Substring removal

Runs only when the last name actually appears as a substring of the first/father name.

Uses `remove_substring()`: pads both strings with spaces, replaces `" {last_name} "` with `" "`, then trims.

```
father_name = "כהן יוסי"
last_name   = "כהן"
→ " כהן יוסי " → replace " כהן " → " יוסי " → "יוסי"
```

If Stage A changes the value → **stop, do not run Stage B**.

If Stage A makes no change (substring not found, or `remove_substring` returns the same string) → fall through to Stage B.

#### Stage B — Positional fallback

Runs **only** when Stage A made no change.

Uses the detected pattern:
- `REMOVE_FIRST`: drop the first token → `" ".join(parts[1:])`
- `REMOVE_LAST`: drop the last token → `" ".join(parts[:-1])`
- `NONE`: no change (Stage B is skipped entirely when pattern is NONE)

Stage B only applies when the field has at least 2 words.

### When Stage B fires without Stage A finding a match

This is the key edge case. If the last name in the source is spelled differently from how it appears in the first/father name field (e.g., different vowel marks, different spacing, or a partial match that `remove_substring` doesn't catch), Stage A will make no change and Stage B will apply positional removal.

Example:
```
last_name   = "כהן"
father_name = "כהן-לוי יוסף"   (after cleaning: "כהן לוי יוסף")
```
Stage A: `"כהן"` IS a substring of `"כהן לוי יוסף"` → removes it → `"לוי יוסף"` → Stage A changed it → stop.

Example where Stage B fires:
```
last_name   = "כהן"
father_name = "כהנים יוסף"   (after cleaning: "כהנים יוסף")
```
Stage A: `"כהן"` is NOT a substring of `"כהנים יוסף"` (word-boundary aware: `" כהן "` not in `" כהנים יוסף "`) → no change.
Stage B (if pattern = REMOVE_FIRST): drops first token → `"יוסף"`.

This means Stage B can remove the wrong word if the pattern was detected on a different set of rows than the current row.

### Single-word names

A single-word first name is never modified (neither Stage A nor Stage B applies). This prevents removing the only word.

### Empty / None values

If either the first/father name or the last name is empty/None after cleaning, no removal is attempted and the cleaned value is returned as-is.

### Examples

| first_name (cleaned) | last_name (cleaned) | Pattern | Stage A result | Stage B result | Final |
|---|---|---|---|---|---|
| `"כהן יוסי"` | `"כהן"` | REMOVE_FIRST | `"יוסי"` (changed) | not run | `"יוסי"` |
| `"יוסי כהן"` | `"כהן"` | REMOVE_LAST | `"יוסי"` (changed) | not run | `"יוסי"` |
| `"יוסי"` | `"כהן"` | REMOVE_FIRST | no change (single word) | not run (single word) | `"יוסי"` |
| `"לוי יוסי"` | `"כהן"` | REMOVE_FIRST | no change (not substring) | `"יוסי"` | `"יוסי"` |
| `"יוסי"` | `"כהן"` | NONE | no change | skipped | `"יוסי"` |
| `"כהן"` | `"כהן"` | REMOVE_FIRST | `""` (empty after removal) | not run | `""` |

---

## 7. Date standardization — Detailed Rules

### Date field types

The system handles two date fields: **birth date** and **entry date**. Each can appear in the source workbook as either:

- **Split columns**: three separate columns for year, month, day (under a parent header like "תאריך לידה" with sub-headers "שנה", "חודש", "יום")
- **Single string column**: one column containing a date string or Excel date value

### Split vs single detection

In the JSON pipeline (`standardizationPipeline`):
- If any of `birth_year`, `birth_month`, `birth_day` keys exist in the row → split path
- If `birth_date` key exists → single path
- Same logic for `entry_*`

Special case in the split path: if only `birth_year` is non-null and `birth_month`/`birth_day` are both null, the year value is treated as a `main_val` (single string) and parsed via the single-value path. This handles merged date cells that openpyxl maps to the first split column.

### Parsing — split columns

`parse_from_split_columns(year_val, month_val, day_val)`:
1. Convert each to `int(float(str(val).strip()))`
2. If year < 100 → expand via two-digit year rule
3. Validate the resulting date

### Parsing — single string / main value

`parse_date_value(raw_value, pattern)` handles these cases in order:

| Input type | Handling |
|---|---|
| `None` or `""` | Returns blank result with status `"תא ריק"` |
| `datetime` or `date` object | Extracts year/month/day directly |
| `int` in range 1–2958465 | Treated as Excel serial date; converted via `openpyxl.utils.datetime.from_excel` |
| String containing month name | `_parse_mixed_month_numeric()` |
| All-digit string | `_parse_numeric_date_string()` |
| ISO-like string `YYYY-MM-DDTHH:MM:SS` | Regex match on first 10 chars |
| String with `/` or `.` separator | `_parse_separated_date_string()` |
| Anything else | Status `"פורמט תאריך לא מזוהה"` |

### Numeric date string parsing

`_parse_numeric_date_string(txt)`:

| Length | Interpretation |
|---|---|
| 8 digits | `DDMMYYYY` |
| 6 digits | `DDMMYYyy` (two-digit year expanded) |
| 4 digits, 1900–2100 | Year only; status `"חסר חודש ויום"` |
| 4 digits, other | `DMYY` (single-digit day and month) |
| Other | Status `"אורך תאריך לא תקין"` |

### Separated date string parsing

`_parse_separated_date_string(txt, pattern)`:
- Normalizes `.` to `/`
- Two-part date (no year): assumes current year
- Three-part date: applies `DateFormatPattern.DDMM` (default) or `MMDD`
- Two-digit year → expanded

**Format pattern detection** (in the direct-Excel processor path only):
- Scans the main date column values
- If first part > 12 and second part ≤ 12 → DDMM vote
- If second part > 12 and first part ≤ 12 → MMDD vote
- MMDD wins only if MMDD votes > DDMM votes; otherwise DDMM

**In the JSON pipeline**, the format pattern is always `DateFormatPattern.DDMM` (hardcoded default). Format detection is not run.

### Two-digit year expansion

```python
current_year = date.today().year   # e.g. 2026
current_two  = current_year % 100  # 26

if yr <= current_two:
    return (current_year // 100) * 100 + yr   # 2000 + yr
else:
    return ((current_year // 100) - 1) * 100 + yr  # 1900 + yr
```

Example (as of 2026): `yr=25` → 2025; `yr=27` → 1927.

### Date validation

`_validate_date(yr, mo, dy)`:
1. Coerce to int
2. Store components on result (even if invalid, so UI can display them)
3. Check day 1–31; if invalid → `"יום לא תקין"`
4. Check month 1–12; if invalid → `"חודש לא תקין"`
5. Check year ≥ 1; if invalid → `"שנה לא תקינה"`
6. Try `datetime(yr, mo, dy)` — catches impossible dates like Feb 30 → `"תאריך לא קיים"`
7. If all pass → `is_valid = True`

### Business rules

Applied after parsing, in `validate_business_rules()`:

| Rule | Condition | Status text |
|---|---|---|
| Empty entry date | `entry_date` with status `"תא ריק"` | Status cleared to `""` (empty entry date is valid) |
| Pre-1900 | `year < 1900` | `"שנה לפני 1900"` |
| Future birth date | `date > today` | `"תאריך לידה עתידי"` |
| Future entry date | `date > today` | `"תאריך כניסה עתידי"` |
| Age > 100 | birth date only | `"גיל מעל 100 (N שנים)"` — note: this is a **warning**, not an error; `is_valid` stays `True` |

### Entry-before-birth cross-validation

In the **direct-Excel processor path**: after both date groups are written, the orchestrator scans each row and appends `"תאריך כניסה לפני תאריך לידה"` to the entry status cell if entry < birth. Cell is formatted pink+bold.

In the **JSON pipeline**: `DateEngine.validate_entry_before_birth()` exists but is NOT called by `standardizationPipeline`. The cross-validation is absent from the web app flow.

### Corrected field writing

**Split date (JSON pipeline):**
- `birth_year_corrected`, `birth_month_corrected`, `birth_day_corrected` — parsed integer values, or original values if parsing failed
- `birth_date_status` — Hebrew status string (empty string if no issue)

**Single date (JSON pipeline):**
- `birth_date_corrected` — formatted as `DD/MM/YYYY` if all components parsed; otherwise original value
- `birth_date_status` — Hebrew status string

**Status cell formatting (direct-Excel path only):**
- `"גיל מעל"` in status → yellow background + bold
- Any other non-empty status → pink background + bold

### Status text reference

| Status | Meaning |
|---|---|
| `""` (empty) | Valid date, no issues |
| `"תא ריק"` | Empty cell (birth date) |
| `"פורמט תאריך לא מזוהה"` | Unrecognized format |
| `"תוכן לא ניתן לפריקה"` | Cannot parse content |
| `"אורך תאריך לא תקין"` | Wrong digit count |
| `"חסר חודש ויום"` | Only year found |
| `"חסר יום"` | Month name found but no day |
| `"תאריך לא ברור"` | Parsing exception |
| `"יום לא תקין"` | Day out of range |
| `"חודש לא תקין"` | Month out of range |
| `"שנה לא תקינה"` | Year < 1 |
| `"תאריך לא קיים"` | Impossible date (e.g. Feb 30) |
| `"שנה לפני 1900"` | Year < 1900 |
| `"תאריך לידה עתידי"` | Birth date in the future |
| `"תאריך כניסה עתידי"` | Entry date in the future |
| `"גיל מעל 100 (N שנים)"` | Age warning (not an error) |
| `"תאריך כניסה לפני תאריך לידה"` | Cross-validation failure (direct-Excel path only) |

### Examples

| Input | Type | Result year/month/day | Status |
|---|---|---|---|
| `1980, 5, 15` | split | 1980/5/15 | `""` |
| `"15/05/1980"` | string | 1980/5/15 | `""` |
| `"15.05.1980"` | string | 1980/5/15 | `""` |
| `"15051980"` | string | 1980/5/15 | `""` |
| `"1997-09-04T00:00:00"` | string | 1997/9/4 | `""` |
| `"05/15/1980"` | string (DDMM default) | invalid (month=15) | `"חודש לא תקין"` |
| `"30/02/2000"` | string | 2000/2/30 stored | `"תאריך לא קיים"` |
| `"1900"` | string | year=1900 | `"חסר חודש ויום"` |
| `""` | string | None | `"תא ריק"` |
| `None` | split | original values | `"תוכן לא ניתן לפריקה"` |
| `datetime(1980,5,15)` | Excel date object | 1980/5/15 | `""` |
| `29221` | Excel serial int | 1980/1/1 | `""` |
| birth date `2030/1/1` | any | 2030/1/1 stored | `"תאריך לידה עתידי"` |
| birth date `1900/1/1` | any | 1900/1/1 | `"גיל מעל 100 (126 שנים)"` |

---

## 8. Identifier standardization — Detailed Rules

### Overview

`IdentifierEngine.normalize_identifiers(id_value, passport_value)` processes both fields together. The two fields interact: an invalid ID may be moved to the passport column.

### Processing order

1. Clean the passport value (`clean_passport`)
2. Treat `"9999"` as no ID (special sentinel value)
3. If no ID → determine status based on passport presence
4. If ID present → run `_process_id_value`

### ID processing (`_process_id_value`)

1. Scan each character: if any non-digit, non-dash character found → move to passport (if passport currently empty)
2. Extract digits only
3. All-zeros → reject (do not move to passport)
4. < 4 digits → move to passport (if empty)
5. > 9 digits → move to passport (if empty)
6. 4–9 digits → pad to 9 with leading zeros → validate checksum

### Israeli ID checksum (Luhn-style)

```
For each digit at position i (0-indexed):
  - Odd positions (i=0,2,4,6,8): multiply by 1
  - Even positions (i=1,3,5,7): multiply by 2; if result > 9, subtract 9
Sum all results. Valid if sum % 10 == 0.
```

### Passport cleaning (`clean_passport`)

Keeps only:
- Digits (0–9)
- ASCII letters (A–Z, a–z)
- Hebrew letters (U+0590–U+05FF, i.e., 1488–1514)
- Dash characters (ASCII hyphen, non-breaking hyphen, figure dash, en-dash, em-dash, horizontal bar, minus sign)

Everything else is dropped.

### Move-to-passport logic

An ID is moved to the passport column when:
- It contains non-digit, non-dash characters (e.g., letters, symbols)
- It has fewer than 4 digits
- It has more than 9 digits

Move only happens if the passport column is currently empty. If passport already has a value, the ID is rejected but NOT moved.

### All-identical-digit rejection

After padding to 9 digits, if all 9 digits are the same (e.g., `"111111111"`, `"999999999"`) → rejected as invalid. This is in addition to the all-zeros check.

### Status text reference

| Status | Condition |
|---|---|
| `"ת.ז. תקינה"` | Valid checksum, no passport |
| `"ת.ז. תקינה + דרכון הוזן"` | Valid checksum + passport present |
| `"ת.ז. לא תקינה"` | Invalid checksum, no passport |
| `"ת.ז. לא תקינה + דרכון הוזן"` | Invalid checksum + passport present |
| `"ת.ז. הועברה לדרכון"` | ID moved to passport (non-digit chars) |
| `"ת.ז. לא תקינה + הועברה לדרכון"` | ID moved to passport (too short/long) |
| `"דרכון הוזן"` | No ID, passport present |
| `"חסר מזהים"` | No ID, no passport |

### Corrected field values

| Scenario | `id_number_corrected` | `passport_corrected` |
|---|---|---|
| Valid ID, no passport | original ID string | `""` |
| Valid ID + passport | original ID string | cleaned passport |
| Invalid checksum | padded 9-digit string | cleaned passport |
| ID moved to passport | `""` | cleaned passport (contains original ID) |
| No ID, passport only | `""` | cleaned passport |
| No ID, no passport | `""` | `""` |

**Note:** When the ID is valid, the output is the **original** string (not the padded form). When invalid (bad checksum), the output is the **padded** 9-digit form.

### Edge cases

| Input | Behavior |
|---|---|
| `id="9999"` | Treated as no ID; status depends on passport |
| `id="000000000"` | All-zeros → rejected; not moved to passport |
| `id="123456789"` (valid checksum) | Kept as-is |
| `id="12345678"` (8 digits) | Padded to `"012345678"`, checksum validated |
| `id="1234567890"` (10 digits) | Moved to passport |
| `id="ABC123"` | Contains letters → moved to passport |
| `id="123-456-789"` | Dashes allowed; digits extracted → `"123456789"` |
| `id=None` | Treated as no ID |
| `passport="AB-123"` | Cleaned: `"AB-123"` (dash kept) |
| `passport="AB 123"` | Cleaned: `"AB123"` (space dropped) |

---

## 9. UI Behavior

### What the user sees

The UI is a single HTML page with these sections:

1. **Upload area**: file input (multi-select), upload button, status text
2. **Session switcher**: one tab per uploaded file; shows `✓` after standardization; hidden when no files uploaded
3. **Sheet selector**: one tab per sheet in the active file
4. **Action bar**: "Run standardization" button, "Export / Download" button, "Delete rows" button (disabled until rows selected)
5. **Grid section**: sheet name title, data grid, row/column count stats, error banner

### Data grid

- HTML `<table>` rendered from the API response
- First column: checkbox for row selection
- Second column: `✕` delete button per row
- Remaining columns: data cells in `display_columns` order

**Column visual classes:**
- Original fields: no special class
- `_corrected` fields: `corrected-cell` (no change) or `corrected-changed` (value differs from original) — highlighted in the UI
- `_status` fields: `status-cell` — distinct visual style

**Editable cells:** All cells except `_status` columns are clickable and become inline text inputs. Clicking a cell that already has an input does nothing (prevents double-edit).

**Edit commit:** On blur or Enter key. On Escape, reverts to original value without saving.

**Edit behavior:** The edit is sent to `PATCH /api/workbook/{id}/sheet/{name}/cell`. On success, the in-memory row is updated and the cell re-renders. If the edited field is a `_corrected` field, the cell class is updated to reflect whether it now differs from the original.

### What is hidden from the user

- `_standardization_failures` key (stripped before display)
- `_standardization_statistics` metadata (not shown in UI)
- Completely empty rows (filtered server-side)
- The numeric helper row (filtered server-side)
- Internal `_serial` key name is shown as-is in the column header (not prettified)

### Row deletion

- Single-row delete: `✕` button on each row
- Multi-row delete: checkboxes + "Delete N rows" button
- Select-all checkbox: selects/deselects all visible rows
- After deletion, the grid re-renders from the updated in-memory data
- The stats line shows "Deleted N row(s). M rows remaining."
- No confirmation dialog

### standardization flow

Clicking "Run standardization":
1. Sends `POST /api/workbook/{id}/normalize` (no `?sheet=` parameter → all sheets)
2. Button shows spinner while waiting
3. On success: session tab gets `✓` badge; current sheet reloads; stats line shows per-sheet results
4. On failure: error banner shown; button re-enabled

### Export flow

Clicking "Export / Download":
1. Sends `POST /api/workbook/{id}/export`
2. Button shows spinner
3. On success: browser downloads the file using the `Content-Disposition` filename
4. The filename is the original stem + `_normalized_YYYYMMDD_HHMMSS.xlsx`

### File switching

Clicking a different file tab:
- Saves the current sheet name to the outgoing session
- Clears the grid and selected rows
- Loads the incoming session's last-viewed sheet (or first sheet if none)
- Does NOT re-normalize or re-extract

### Awkward / incomplete parts

- **Bulk export buttons are commented out.** The code for `exportBulk()` and `exportSelected()` exists in `app.js` but the buttons that call them are wrapped in a `/* ... */` comment block. Users cannot export multiple files as ZIP from the UI.
- **No undo.** Deleted rows are gone for the session. Edits cannot be reverted except by re-standardizing (which re-extracts from disk, discarding manual edits).
- **Re-standardizing discards manual edits.** standardization re-extracts from the working copy on disk and replaces the in-memory dataset. Any manual cell edits made before standardization are lost.
- **No pagination.** Large sheets (thousands of rows) render all rows at once, which can be slow.
- **Status columns not editable.** This is intentional but there is no visual indicator explaining why clicking a status cell does nothing.
- **`_serial` column header.** The synthetic serial number column shows `_serial` as its header text, which is an internal key name, not a user-friendly label.
- **Edit validation.** The edit API accepts any string value for any field. There is no client-side or server-side validation of the new value's format.

---

## 10. Export Behavior

### Export schema

The web app export (`ExportService`) creates a new workbook with one sheet per source sheet. The column schema is fixed and determined by the canonical sheet name:

**DayarimYahidim** (14 columns):
```
MosadID, SugMosad, ShemPrati, ShemMishpaha, ShemHaAv,
MisparZehut, Darkon, Min,
ShnatLida, HodeshLida, YomLida,
ShnatKnisa, HodeshKnisa, YomKnisa
```

**MeshkeyBayt / AnasheyTzevet** (15 columns):
```
MosadID, SugMosad, MisparDiraBeMosad, ShemPrati, ShemMishpaha, ShemHaAv,
MisparZehut, Darkon, Min,
ShnatLida, HodeshLida, YomLida,
ShnatKnisa, HodeshKnisa, YomKnisa
```

**Unmatched sheets**: fall back to the DayarimYahidim 14-column schema.

### Sheet name canonicalization

Source sheet names are matched against keyword patterns (after NFC standardization + whitespace collapse):

| Source contains | Canonical name |
|---|---|
| `"דיירים יחידים"` or `"דיירים"` | `DayarimYahidim` |
| `"מתגוררים במשקי בית"`, `"משקי בית"`, or `"מתגוררים"` | `MeshkeyBayt` |
| `"אנשי צוות ובני משפחותיהם"`, `"אנשי צוות"`, or `"צוות"` | `AnasheyTzevet` |
| Anything else | original sheet name (unchanged) |

### Field mapping (export header → JSON key)

| Export header | Source JSON key |
|---|---|
| `MosadID` | `MosadID` (from derived columns) |
| `SugMosad` | `SugMosad` (always blank — not populated) |
| `MisparDiraBeMosad` | `MisparDiraBeMosad` (from source row if present) |
| `ShemPrati` | `first_name_corrected` |
| `ShemMishpaha` | `last_name_corrected` |
| `ShemHaAv` | `father_name_corrected` |
| `MisparZehut` | `id_number_corrected` |
| `Darkon` | `passport_corrected` |
| `Min` | `gender_corrected` |
| `ShnatLida` | `birth_year_corrected` |
| `HodeshLida` | `birth_month_corrected` |
| `YomLida` | `birth_day_corrected` |
| `ShnatKnisa` | `entry_year_corrected` |
| `HodeshKnisa` | `entry_month_corrected` |
| `YomKnisa` | `entry_day_corrected` |

**Corrected-only rule:** There is no fallback to original values. If `first_name_corrected` is absent or empty, `ShemPrati` is blank in the export. This means exporting before standardization produces a mostly-blank file.

### Row filtering in export

The same two filters applied in the UI are applied in export:

1. **Empty rows**: rows where every original-column cell is None/empty are dropped
2. **Numeric helper row**: if the first row has all-numeric values in original columns, it is dropped

Row validity for export: a row is included if it passes the above filters. There is no additional "at least one personal field non-empty" check in the web app export (unlike the VBA-parity `ExportEngine` which checks `ShemPrati`, `ShemMishpaha`, etc.).

### RTL and formatting

- Each export sheet has `ws.sheet_view.rightToLeft = True`
- Header cells have `Alignment(horizontal="right")`
- No other formatting (no bold headers, no column widths set)

### Output file location and naming

- Saved to `output/` directory (or `%LOCALAPPDATA%\Excelstandardization\output\` in packaged mode)
- Filename: `{original_stem}_normalized_{YYYYMMDD_HHMMSS}.xlsx`
- Always `.xlsx` regardless of source format
- No collision handling (timestamp makes each export unique)

### Export without prior standardization

If the user clicks Export without first running standardization:
- `ExportService` auto-extracts the workbook from disk
- No `_corrected` fields exist in the rows
- All personal data columns (`ShemPrati`, etc.) will be blank in the export
- `MosadID` may still be populated if the scanner found it

### Known limitations

- `SugMosad` is always blank. The field exists in the schema but there is no source for it.
- `MisparDiraBeMosad` is only populated if the source row already has a `MisparDiraBeMosad` key. The system does not extract this from the source workbook.
- The export does not include status columns, original values, or any diagnostic information.
- The serial number column is not included in the export schema.
- No cell-level formatting (no pink highlights, no bold) in the export output.

---

## 11. Installer and Launcher Behavior

### Launch flow

`launcher.py` is the PyInstaller entry point:

1. Sets up logging (console + file)
2. Finds a free port starting at 8765 (tries 8765–8864, then any free port)
3. Detects Chrome installation
4. Starts a background thread that sleeps 1.5 seconds then opens the browser
5. Starts Uvicorn on `127.0.0.1:{port}` (blocking call)

The server binds to `127.0.0.1` only — no network exposure.

### Browser behavior

Chrome detection checks in order:
1. `C:\Program Files\Google\Chrome\Application\chrome.exe`
2. `C:\Program Files (x86)\Google\Chrome\Application\chrome.exe`
3. `%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe`
4. `where chrome` command (registry fallback)

If Chrome is found: opens with `--app={url} --new-window` (app mode, no address bar).
If Chrome is not found: falls back to `webbrowser.open(url)` (system default browser).

### Local data directories

In packaged mode (`sys.frozen = True`):
- `%LOCALAPPDATA%\Excelstandardization\uploads\`
- `%LOCALAPPDATA%\Excelstandardization\work\`
- `%LOCALAPPDATA%\Excelstandardization\output\`
- `%LOCALAPPDATA%\Excelstandardization\app.log`

In development mode: `uploads/`, `work/`, `output/` relative to the working directory.

The Inno Setup installer pre-creates these directories with `users-full` permissions so the app can write without admin rights at runtime.

### Logging

Two handlers:
- `StreamHandler(sys.stdout)` — INFO level
- `FileHandler(app.log)` — INFO level (not DEBUG)

Log format: `%(asctime)s [%(levelname)s] %(name)s: %(message)s`

The log file is not rotated and grows indefinitely.

### Packaging

Built with PyInstaller (`Excelstandardization.spec`). One-folder bundle (not one-file). The `dist/Excelstandardization/` folder contains the executable and all dependencies.

Asset path resolution: in the bundle, `sys._MEIPASS` points to the extraction directory. `webapp/app.py` uses `Path(sys._MEIPASS) / "webapp"` to find static files and templates.

### Installer (Inno Setup)

- Version: 1.0.1
- Requires admin for installation (installs to `Program Files`)
- x64 Windows 10+ only
- Creates Start Menu and optional desktop shortcut
- Offers to launch after installation
- Uninstall removes the app directory but leaves user data in `%LOCALAPPDATA%\Excelstandardization`

### Packaging limitations

- Sessions are lost on restart (in-memory only)
- No auto-update mechanism
- No crash reporting
- The log file is not accessible from the UI
- If port 8765 is in use, the app silently picks another port; the user sees the correct URL in the console but not in the browser title bar (app mode hides the address bar)
- If Chrome is not installed, the app opens in the default browser which may not support all features (though the app uses only standard HTML/CSS/JS)

---

## 12. Edge-Case Catalogue

### Sheet loading edge cases

| Case | Current behavior |
|---|---|
| Sheet with no recognizable headers | Skipped; `SheetDataset` returned with `skipped=True`; not shown in UI |
| Sheet where best header row scores < 3 | Same as above |
| Sheet where the header row is row 1 and there is no data | `data_start_row = 2`; `end_row = 1`; zero rows extracted |
| Sheet with merged cells spanning the header row | Merged cell value read from top-left cell; all spanned columns marked as processed |
| Sheet with a column-index row (1, 2, 3…) after headers | Row detected and skipped at extraction time AND at display time |
| Sheet where the column-index row has gaps (e.g., 1, 2, 4) | NOT detected as column-index row (consecutive check fails); treated as data |
| Sheet with headers in row 25+ | Not detected (max_scan_rows=30 default; would be found if within 30 rows) |
| Sheet with two date groups of the same type (two birth date columns) | Both groups detected and processed in the direct-Excel path; JSON pipeline only uses the first group |
| Workbook with no valid sheets | `WorkbookDataset.sheets = []`; standardization raises HTTP 500 |
| `.xlsm` file | Accepted; VBA macros are not executed; `keep_vba=True` only in direct-Excel path |

### Name standardization edge cases

| Case | Current behavior |
|---|---|
| Name is only digits (`"123"`) | All digits dropped → `""` |
| Name is only symbols (`"---"`) | Hyphens → spaces → collapsed → `""` |
| Name is only a title (`"ד\"ר"`) | After char filtering: `"דר"` → removed as unwanted token → `""` |
| Name has mixed Hebrew and English (`"יוסי Smith"`) | Language = HEBREW (Hebrew ≥ English); English letters dropped → `"יוסי"` |
| Name has equal Hebrew and English letters | Language = HEBREW (tie goes to Hebrew); English dropped |
| Name is all English with Hebrew title (`"Dr. יוסי"`) | Language = HEBREW (Hebrew letters present); English dropped → `"יוסי"` |
| Name with geresh (`"ג'ורג'"`) | Geresh (U+05F3) is not in Hebrew letter range → dropped → `"גורג"` |
| Name with gershayim (`"צ\"ל"`) | Gershayim (U+05F4) dropped → `"צל"` |
| Name with zero-width chars | Stripped in step 1 |
| Name with Arabic-Indic digits (`"١٢٣"`) | Converted to `"123"` then dropped |
| `None` input | Returns `""` |
| Empty string input | Returns `""` |
| Very long name | No length limit; processed in full |

### Last-name removal edge cases

| Case | Current behavior |
|---|---|
| Last name is a substring of a word in first name (not a whole word) | Stage A: `remove_substring` uses space-padded replacement, so `"כהנים"` is NOT matched by `"כהן"` → Stage B fires |
| First name is a single word | Neither Stage A nor Stage B modifies it |
| Last name equals entire first name | Stage A removes it → `""` |
| Pattern detected as NONE | No removal at all |
| Sample size < 3 rows with both fields | Pattern = NONE |
| Last name appears in 2 of 5 sample rows | `contain < 3` → NONE |
| Pattern detected as REMOVE_FIRST but current row has last name at end | Stage A removes it (substring match); Stage B not reached |
| Pattern detected as REMOVE_FIRST but last name not in current row | Stage A no-op; Stage B removes first word regardless |

### Date edge cases

| Case | Current behavior |
|---|---|
| Split date where year column contains a full date string | Detected: if month and day are null, year value treated as main_val |
| Split date where year column contains a `datetime` object | Detected: treated as main_val |
| Date `"30/02/2000"` | Parsed: year=2000, month=2, day=30 stored; status `"תאריך לא קיים"` |
| Date `"00/00/0000"` | year=0, month=0, day=0; status `"יום לא תקין"` |
| Date `"1980"` (4-digit year) | Parsed as year-only; status `"חסר חודש ויום"` |
| Date `"15/5"` (two-part) | Assumes current year; parsed as DD/MM/current_year |
| Date `"January 15 1980"` | Month name detected; parsed correctly |
| Date `"15 ינואר 1980"` | Hebrew month name detected; parsed correctly |
| Two-digit year `"25"` (as of 2026) | Expanded to 2025 |
| Two-digit year `"27"` (as of 2026) | Expanded to 1927 |
| Empty entry date | Status cleared to `""` (valid; entry date is optional) |
| Entry date before birth date | Direct-Excel path: warning appended to entry status. JSON pipeline: not checked. |
| Age exactly 100 | No warning (rule is `> 100`) |
| Age 101 | Status `"גיל מעל 100 (101 שנים)"` — `is_valid` stays `True` |

### Identifier edge cases

| Case | Current behavior |
|---|---|
| ID `"9999"` | Treated as no ID |
| ID `"000000000"` | All-zeros → rejected; not moved to passport |
| ID `"111111111"` | All-identical → rejected; not moved to passport |
| ID with dashes (`"123-456-789"`) | Dashes allowed; digits extracted → `"123456789"` |
| ID with letters (`"AB123456"`) | Contains non-digit → moved to passport (if passport empty) |
| ID with letters, passport already has value | Moved-to-passport logic skipped; ID rejected but passport unchanged |
| ID `"12345"` (5 digits) | Padded to `"000012345"`; checksum validated |
| ID `"1234567890"` (10 digits) | Too long → moved to passport |
| ID `"123"` (3 digits) | Too short → moved to passport |
| Passport with spaces | Spaces dropped |
| Passport with Hebrew letters | Kept |
| Passport `None` | Treated as empty |

### UI edge cases

| Case | Current behavior |
|---|---|
| Clicking "Export" before "Normalize" | Export runs; corrected fields absent → mostly blank output |
| Clicking "Normalize" twice | Second standardization re-extracts from disk; manual edits from first session are lost |
| Deleting all rows | Sheet shows "No data rows found" message |
| Editing a `_corrected` field | Allowed; cell class updates to reflect new vs original comparison |
| Editing an original field | Allowed; does NOT re-run standardization; corrected field retains old value |
| Uploading the same file twice | Two separate sessions created; independent state |
| Uploading a file with no recognizable sheets | Upload succeeds (sheet names returned); loading any sheet shows "No data rows found" |
| Very large file (many rows) | No pagination; browser may be slow rendering the grid |

---

## 13. Gaps, Limitations, and Open Questions

### Critical gaps

**1. Manual edits are discarded on re-standardization**
`standardizationService.normalize()` re-extracts from disk and replaces the in-memory dataset. Any manual cell edits made before standardization are silently lost. The `record.edits` dict is populated but never replayed. This is a significant usability problem: the intended workflow is "normalize → review → edit → export", but re-standardizing after editing destroys the edits.

**2. Entry-before-birth cross-validation missing from JSON pipeline**
`DateEngine.validate_entry_before_birth()` exists but is not called by `standardizationPipeline`. The cross-validation only runs in the legacy direct-Excel processor path. The web app never flags entry dates that precede birth dates.

**3. `SugMosad` always blank**
The export schema includes `SugMosad` but there is no source for this value. It is always empty in the output. The business meaning and source of this field are unclear.

**4. `MisparDiraBeMosad` not extracted**
The field is in the export schema for MeshkeyBayt and AnasheyTzevet sheets, but the extraction pipeline does not look for it in the source workbook. It will only appear in the export if it was already present as a key in the source rows (which it won't be from a fresh extraction).

**5. Session data lost on restart**
All sessions, uploaded files, and standardization results are in-memory only. A server restart (or crash) loses everything. Users must re-upload and re-normalize.

### Behavioral inconsistencies

**6. Date format detection not used in JSON pipeline**
The direct-Excel processor detects DDMM vs MMDD format from the data. The JSON pipeline always uses DDMM. This means ambiguous dates (e.g., `"05/06/1980"`) may be interpreted differently depending on which path is used.

**7. Pattern detection sample size**
The last-name removal pattern is detected from the first 10 rows that have both fields populated. If the first 10 rows are atypical (e.g., all have the last name in the first name, but the rest don't), the pattern will be applied incorrectly to all rows.

**8. `_serial` column header**
The synthetic serial number column uses `_serial` as its key and header. This is an internal name, not a user-friendly label. It should be displayed as something like `"#"` or `"מספר שורה"`.

**9. Highlight comparison in direct-Excel path**
`ExcelWriter.highlight_changed_cells()` normalizes numeric representations (`"1.0"` == `"1"`) before comparing. The JSON pipeline does not apply this standardization when deciding whether to show a cell as `corrected-changed` in the UI — it uses a simple `value !== origVal` JavaScript comparison.

**10. Two-row header detection is heuristic**
The scoring system for detecting whether a sheet has a two-row header can misfire on unusual layouts. There is no fallback or user override.

### Missing features

**11. No undo / redo**
Deleted rows and edited cells cannot be reverted except by re-uploading the file.

**12. No pagination**
Large sheets render all rows at once. This can cause browser performance issues for sheets with thousands of rows.

**13. Bulk ZIP export UI disabled**
The `exportBulk()` and `exportSelected()` functions exist in `app.js` but the buttons are commented out. Users cannot export multiple files at once from the UI.

**14. No session persistence**
There is no way to resume a session after closing the browser or restarting the server.

**15. No progress indicator for large files**
standardization of large workbooks can take several seconds. The button shows a spinner but there is no per-sheet progress.

**16. No validation of edited values**
The edit API accepts any string. Editing a date field to `"abc"` is accepted without error. The corrected value will be wrong but no warning is shown.

### Ambiguous / needs business validation

**17. Age > 100 is a warning, not an error**
`is_valid` stays `True` for ages over 100. The status text is written but the corrected date values are still exported. Is this the intended behavior, or should these rows be flagged more prominently?

**18. All-identical-digit ID rejection**
IDs like `"111111111"` are rejected as invalid. Is this a business rule or a heuristic? The VBA source is the reference but this behavior should be confirmed.

**19. `"9999"` as no-ID sentinel**
The value `"9999"` is treated as "no ID provided". This is presumably a legacy convention from the source data. It should be confirmed whether other sentinel values exist.

**20. Empty entry date is valid**
An empty entry date clears the status to `""` and is treated as valid. Is this correct for all sheet types, or only for some?

**21. Pattern detection threshold of 3**
The last-name removal pattern requires at least 3 out of 5 sample rows to match. This threshold is hardcoded. It may be too low (false positives) or too high (false negatives) depending on data quality.

**22. `MosadID` scanning scope**
The MosadID scanner looks for a label/value pair anywhere in the worksheet, including outside the main data table. If a sheet has multiple label/value pairs matching the pattern, only the first one found (top-to-bottom, left-to-right) is used. This may not always be the correct one.

**23. Export without standardization**
Exporting before standardization produces a file with blank personal data columns. Should the system warn the user, or refuse to export, or fall back to original values?

---

*Document generated from codebase analysis. Last updated: April 2026.*
