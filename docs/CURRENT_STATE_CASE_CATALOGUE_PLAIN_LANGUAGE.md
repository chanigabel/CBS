# Current-State Case Catalogue — Plain Language Edition

> **Who this is for:** Business reviewers, analysts, QA testers, and project owners who need to understand exactly what the system does — without reading code.
>
> **What this is:** A complete description of every meaningful behavior the system currently implements, derived entirely from the real codebase. Nothing is invented. Nothing is aspirational. If the system does something unexpected, it is documented here.
>
> **How to read it:** Each section covers one area of the system. Each case has a plain-language explanation of the input, what the system does, and what comes out. Case IDs match the technical catalogue exactly so developers can cross-reference.

---

## Table of Contents

1. [Sheet Loading](#1-sheet-loading)
2. [Header Detection](#2-header-detection)
3. [Row Filtering](#3-row-filtering)
4. [Column Ordering in the Display](#4-column-ordering-in-the-display)
5. [Name standardization](#5-name-standardization)
6. [Last-Name Removal from the First Name Field](#6-last-name-removal-from-the-first-name-field)
7. [Last-Name Removal from the Father Name Field](#7-last-name-removal-from-the-father-name-field)
8. [Gender standardization](#8-gender-standardization)
9. [Date Parsing and Statuses](#9-date-parsing-and-statuses)
10. [Identifier and Passport Logic](#10-identifier-and-passport-logic)
11. [Institution ID (MosadID) Extraction](#11-institution-id-mosadid-extraction)
12. [Session: Editing, Deleting, and Re-standardizing](#12-session-editing-deleting-and-re-standardizing)
13. [Export Behavior](#13-export-behavior)
14. [Differences Between the Web App and the Command-Line Tool](#14-differences-between-the-web-app-and-the-command-line-tool)
15. [Observed Risky Behaviors](#15-observed-risky-behaviors)

---

## 1. Sheet Loading

**What this area does:**
When a user uploads an Excel file, the system must figure out where the actual data table is on each sheet. Excel files often have titles, logos, or blank rows above the real data. The system scans the first 30 rows of each sheet, scores each row based on how much it looks like a header row (does it contain recognizable field names? is it mostly text rather than numbers?), and picks the best candidate. If no row scores high enough, the sheet is skipped entirely.

The system also handles two special layout patterns common in Israeli institutional forms:
- **Two-row headers:** A parent header row (e.g., "Birth Date") with a sub-row beneath it containing "Year", "Month", "Day".
- **Column-index rows:** Some forms include a row of sequential numbers (1, 2, 3…) immediately after the headers to label column positions. The system detects and skips these.

**Source:** `ExcelReader.detect_table_region()` in `excel_reader.py`

---

### Cases

---

**SL-01 — Standard layout: headers in row 1**

- **Input:** A sheet where the first row contains column headers and data starts in row 2.
- **What the system does:** Scores row 1 highest; identifies it as the header row.
- **Output:** Data extraction begins from row 2.
- **Why this matters:** This is the most common layout. It must work correctly every time.

---

**SL-02 — Title rows above the headers**

- **Input:** A sheet where rows 1–2 contain a title or logo, and the actual column headers are in row 3.
- **What the system does:** Scores all rows; row 3 scores highest because it contains recognizable field names.
- **Output:** Data extraction begins from row 4.
- **Why this matters:** Many institutional forms have decorative or informational content above the data table. The system handles this automatically.

---

**SL-03 — Two-row header (date sub-headers)**

- **Input:** A sheet where one row contains "Birth Date" as a merged header, and the row below it contains "Year", "Month", "Day" as sub-columns.
- **What the system does:** Detects the parent-child relationship; sets `header_rows=2`; data starts two rows below the parent header.
- **Output:** The three date columns are correctly identified as year, month, and day of birth.
- **Why this matters:** Split date columns are the standard format in these forms. Misidentifying the header row would cause all date data to be lost.

---

**SL-04 — Sheet with no recognizable headers**

- **Input:** A sheet where no row scores high enough to be a header (score below 3).
- **What the system does:** Marks the sheet as skipped.
- **Output:** The sheet does not appear in the data grid or export.
- **Why this matters:** Summary sheets, charts, or instruction sheets in the same workbook should not be processed as data.

---

**SL-05 — Completely empty sheet**

- **Input:** A sheet with no cells containing any data.
- **What the system does:** No rows to score; returns no table region.
- **Output:** Sheet skipped.
- **Why this matters:** Empty sheets should not cause errors.

---

**SL-06 — Column-index helper row (1, 2, 3…)**

- **Input:** A sheet where the row immediately after the headers contains sequential integers (1, 2, 3, 4…) labeling column positions.
- **What the system does:** Detects that this row is all small consecutive integers; skips it.
- **Output:** Data extraction begins from the row after the index row.
- **Why this matters:** These helper rows appear in many standard forms. Without this detection, the numbers would appear as the first data row.

---

**SL-07 — Column-index row with a gap (e.g., 1, 2, 4)**

- **Input:** A row of integers after the headers, but with a gap in the sequence (e.g., 1, 2, 4 — missing 3).
- **What the system does:** The consecutive-sequence check fails; the row is NOT treated as a helper row.
- **Output:** The row is treated as a normal data row and included in the data.
- **Why this matters:** The detection is strict. A gap means the row might be real data, not a column index.

---

**SL-08 — Column-index row with only 2 values**

- **Input:** A row after the headers with only 2 numeric values.
- **What the system does:** Requires at least 3 values to be confident; does not skip the row.
- **Output:** Row treated as data.
- **Why this matters:** Two numbers could easily be real data (e.g., an ID and a year). The system avoids false positives.

---

**SL-09 — Table ends with 5 consecutive empty rows**

- **Input:** A data table followed by 5 or more completely empty rows.
- **What the system does:** Stops scanning after 5 consecutive empty rows; marks the last row with data as the table end.
- **Output:** Only rows up to the last data row are extracted.
- **Why this matters:** Prevents the system from scanning thousands of empty rows in large workbooks.

---

**SL-10 — 4 empty rows then more data**

- **Input:** A table with a 4-row gap in the middle, then more data rows.
- **What the system does:** Does not stop at 4 empty rows; continues scanning.
- **Output:** All data rows, including those after the gap, are extracted.
- **Why this matters:** Some forms have intentional gaps. The threshold is 5, not 4.

---

**SL-11 — Merged cells in the header row**

- **Input:** A header row where some cells are merged (e.g., "Birth Date" spans three columns).
- **What the system does:** Reads the value from the top-left cell of the merged range; marks all spanned columns as processed.
- **Output:** The merged header is correctly identified once; no duplicate columns.
- **Why this matters:** Merged headers are standard in these forms for date groups.

---

**SL-12 — Header column already marked as "corrected"**

- **Input:** A column whose header contains the word "מתוקן" (corrected) or "corrected".
- **What the system does:** Ignores this column entirely.
- **Output:** The column does not appear in the extracted data.
- **Why this matters:** If a file has already been processed once, the corrected columns from the previous run should not be re-processed as source data.

---

**SL-13 — Column with an unrecognized header**

- **Input:** A column whose header does not match any known field name (e.g., "הערות" / "Notes").
- **What the system does:** Includes the column using a sanitized version of the header text as the field name.
- **Output:** The column appears in the data grid with its original header as the column name.
- **Why this matters:** No source data should be silently dropped. Unknown columns are passed through as-is.

---

**SL-14 — Name fields only on the sub-header row**

- **Input:** A two-row header layout where "First Name", "Last Name" appear only on the second header row (not the parent row).
- **What the system does:** After processing the parent row, scans the sub-header row for any columns not yet mapped.
- **Output:** Name fields are correctly identified.
- **Why this matters:** In some form layouts, name columns sit alongside date sub-columns on the same sub-header row.

---

**SL-15 — Headers beyond row 30**

- **Input:** A sheet where the actual column headers are in row 31 or later.
- **What the system does:** Only scans the first 30 rows; does not find the headers.
- **Output:** Sheet is skipped.
- **Why this matters:** This is a known limitation. Forms with more than 30 rows of preamble will not be processed.

---

**SL-16 — Macro-enabled workbook (.xlsm)**

- **Input:** An `.xlsm` file (Excel with macros).
- **What the system does:** Opens the file and reads data normally. The macros are not executed.
- **Output:** Data extracted as normal.
- **Why this matters:** Many institutional forms use `.xlsm` format. The system handles them without running any embedded code.

---
