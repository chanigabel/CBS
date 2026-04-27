# Row Identity Stable Edits — Bugfix Design

## Overview

The pipeline currently uses mutable array indices as row identity throughout every
stage: extraction, normalization, display, editing, and export.  When any row is
removed or skipped (empty-row filter, numeric helper-row filter, user deletion, or
re-normalization reorder), all subsequent array indices shift, causing edits to be
written to the wrong source rows on export.

The fix introduces a stable, source-anchored row identity at the earliest possible
point — the Excel extraction layer — and threads it through every subsequent stage
without alteration.  No normalization logic changes; only identity bookkeeping and
the lookup strategy for edits/deletes change.

---

## Glossary

- **Bug_Condition (C)**: The condition that triggers the bug — any pipeline state
  where a row's array index differs from the index that was used when an edit was
  recorded, OR where the export iterates rows by sequential output position rather
  than by `_excel_row_number`.
- **Property (P)**: The desired behavior — every edit is permanently bound to the
  exact source row via `_row_uid`, and the export writes corrected values to the
  correct source row regardless of how many rows were skipped or deleted.
- **Preservation**: All existing normalization output, display filtering, export
  schema, and column ordering must remain unchanged by this fix.
- **`_row_uid`**: Stable string key `"{sheet_name}:{excel_row_number}"` assigned at
  extraction time and never mutated.
- **`_excel_row_number`**: 1-based Excel row number of the data row (not the header).
- **`_source_data_index`**: 0-based position in the raw extracted array before any
  filtering.
- **`_source_sheet`**: Worksheet title at extraction time (redundant with the sheet
  key in `_row_uid` but kept for debugging convenience).
- **`isBugCondition`**: Pseudocode predicate that identifies a buggy pipeline state.
- **`ExcelToJsonExtractor.extract_sheet_to_json`**: In
  `src/excel_normalization/io_layer/excel_to_json_extractor.py` — builds the raw
  `JsonRow` list from an openpyxl worksheet.
- **`NormalizationPipeline.normalize_dataset`**: In
  `src/excel_normalization/processing/normalization_pipeline.py` — copies rows and
  applies all engines; must propagate identity fields.
- **`EditService.edit_cell` / `delete_rows`**: In `webapp/services/edit_service.py`
  — currently looks up rows by array index; must switch to `_row_uid` lookup.
- **`NormalizationService.normalize`**: In `webapp/services/normalization_service.py`
  — replays edits after re-normalization; must match by `_row_uid`.
- **`WorkbookService.get_sheet_data`**: In `webapp/services/workbook_service.py`
  — strips underscore-prefixed keys before sending to the UI; must pass `_row_uid`
  and `_excel_row_number` through.
- **`ExportService.visible_rows`**: In `webapp/services/export_service.py` — filters
  rows for export; must preserve identity fields.
- **`SessionRecord.edits`**: In `webapp/models/session.py` — dict keyed by
  `(sheet_name, row_idx, field_name)`; must be rekeyed to
  `(sheet_name, row_uid, field_name)`.
- **`renderGrid` / `makeEditable` / `_deleteRows`**: In `webapp/static/app.js` —
  currently send `rowIndex`; must send `row_uid`.

---

## Bug Details

### Bug Condition

The bug manifests whenever the array index of a row in `sheet.rows` differs from
the index that was used when an edit was recorded, OR when the export iterates rows
by sequential output position rather than by `_excel_row_number`.  This happens
whenever any row is removed or skipped at any pipeline stage.

**Formal Specification:**
```
FUNCTION isBugCondition(pipeline_state)
  INPUT: pipeline_state describing rows in memory and the edit store
  OUTPUT: boolean

  RETURN (
    EXISTS row IN sheet.rows WHERE
      row does NOT have a stable _row_uid field
    OR
    EXISTS edit IN record.edits WHERE
      edit is keyed by (sheet_name, array_index, field_name) AND
      array_index no longer points to the same source row
    OR
    export writes row at output_position WITHOUT referencing _excel_row_number
  )
END FUNCTION
```

### Examples

- **Skipped helper row**: Excel rows 1–header, 2–column-index helper, 3–"Yotam",
  4–"Rachel".  After the helper-row filter, "Yotam" is at array index 0 and
  "Rachel" at index 1.  The UI sends `rowIndex=1` for "Rachel".  On export the
  system writes the edit to output row 2 (0-based index 1), which maps back to
  "Yotam" in the source — data corruption.

- **Empty row removal**: Excel rows 3 and 5 are empty; row 4 contains "David".
  After empty-row filtering, "David" is at array index 0.  An edit stored as
  `(sheet, 0, "first_name_corrected")` is replayed after re-normalization and
  hits whatever row happens to be at index 0 in the new array — which may be a
  different person if rows were reordered.

- **Re-normalization after deletion**: User deletes row at index 2, then
  re-normalizes.  The edit store still contains keys with the old indices.
  Replay writes edits to wrong rows.

- **No skips (non-buggy)**: All rows present, no filtering, no deletion.
  Array indices happen to match Excel row numbers minus the header offset.
  Edits land on the correct rows by coincidence — this is the preserved path.

---

## Expected Behavior

### Preservation Requirements

**Unchanged Behaviors:**
- Normalization engines (`NameEngine`, `GenderEngine`, `DateEngine`,
  `IdentifierEngine`) must produce identical corrected field values for every row.
- The UI grid must continue to hide all underscore-prefixed internal keys
  (`_normalization_failures`, `_birth_year_auto_completed`, etc.) — `_row_uid` and
  `_excel_row_number` are also underscore-prefixed and must NOT appear as visible
  columns.
- The empty-row filter and numeric helper-row filter must continue to hide those
  rows from the UI and from export output.
- The export schema (column order, sheet-name mapping, right-to-left direction,
  `EXPORT_MAPPING`) must remain unchanged.
- Serial-number and MosadID derived-column injection must continue to work.
- Multi-sheet workbooks must continue to process each sheet independently.
- Mouse clicks, keyboard shortcuts, and all other UI interactions unrelated to
  row identity must be unaffected.

**Scope:**
All inputs that do NOT involve row removal, skipping, or re-normalization after
deletion should produce identical output before and after the fix.  This includes:
- Workbooks where every extracted row survives all filters unchanged.
- Edit-then-export flows with no intervening row removal.
- Normalization runs on sheets with no prior manual edits.

---

## Hypothesized Root Cause

The root cause is confirmed (per the requirements document and user analysis):

1. **No stable identity at extraction** (`ExcelToJsonExtractor.extract_sheet_to_json`):
   Rows are appended to a plain list with no `_row_uid` or `_excel_row_number`
   field.  The only available identity after extraction is the list index.

2. **Index-shifting filters** (`WorkbookService.get_sheet_data`,
   `ExportService.visible_rows`): Both apply empty-row and helper-row filters that
   remove elements from the list, shifting all subsequent indices.  The filtered
   view index sent by the UI is therefore not the same as the backing-dataset index.

3. **Index-based edit storage** (`EditService.edit_cell`, `SessionRecord.edits`):
   Edits are stored as `(sheet_name, row_idx, field_name)` where `row_idx` is the
   filtered-view index from the UI — not even the backing-dataset index.

4. **Index-based edit replay** (`NormalizationService.normalize`): After
   re-normalization, edits are replayed by `row_idx` against the new array, which
   may have a different length and order.

5. **Sequential export output** (`ExportService.export` loop): The export loop
   writes rows to sequential output positions (`out_row += 1`) with no reference
   to `_excel_row_number`, so any row removal before export shifts all subsequent
   output rows.

6. **Identity not propagated through normalization**
   (`NormalizationPipeline.normalize_dataset`): `normalize_row` does a shallow
   `copy()` of the input row dict, which would preserve any `_row_uid` field —
   but since extraction never adds it, there is nothing to preserve.

---

## Correctness Properties

Property 1: Bug Condition — Stable Row Identity Through Edit and Export

_For any_ pipeline state where one or more rows are skipped or removed before the
edited row (i.e. `isBugCondition` returns true), the fixed pipeline SHALL store
the edit keyed by `_row_uid` and the export SHALL write the corrected value to the
exact Excel row identified by `_excel_row_number`, leaving all other rows
unchanged.

**Validates: Requirements 2.1, 2.2, 2.3, 2.4, 2.5, 2.6**

Property 2: Preservation — Identical Output When No Rows Are Skipped

_For any_ pipeline state where no rows are skipped, filtered, or deleted (i.e.
`isBugCondition` returns false), the fixed pipeline SHALL produce the same
normalized field values, the same display grid, and the same exported workbook as
the original pipeline, preserving all existing normalization, display, and export
behavior.

**Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7**

---

## Fix Implementation

### Changes Required

Assuming the confirmed root cause analysis:

---

**File**: `src/excel_normalization/io_layer/excel_to_json_extractor.py`

**Function**: `extract_sheet_to_json`

**Specific Changes**:
1. **Inject identity fields before appending each row**: Inside the
   `for row_num in range(...)` loop, immediately after `json_row` is built by
   `extract_row_to_json`, add:
   ```python
   json_row["_source_sheet"] = worksheet.title
   json_row["_excel_row_number"] = row_num          # actual 1-based Excel row
   json_row["_source_data_index"] = len(rows)       # 0-based position before append
   json_row["_row_uid"] = f"{worksheet.title}:{row_num}"
   ```
   These four fields are added **before** the `skip_empty_rows` check so they are
   present on every row that enters the list, including rows that will later be
   filtered by the service layer.

---

**File**: `src/excel_normalization/processing/normalization_pipeline.py`

**Function**: `normalize_row`

**Specific Changes**:
1. **Propagate identity fields through normalization**: After `result = json_row.copy()`,
   the four identity fields (`_source_sheet`, `_excel_row_number`,
   `_source_data_index`, `_row_uid`) are already present in the copy because
   `dict.copy()` is a shallow copy.  No additional code is needed here — but the
   `_apply_birth_year_majority_correction` method must NOT strip these fields.
   Verify that only `_birth_year_auto_completed` and `_entry_year_auto_completed`
   are popped; the identity fields must survive.

---

**File**: `webapp/models/session.py`

**Class**: `SessionRecord`

**Specific Changes**:
1. **Update edits dict key type comment**: Change the docstring from
   `{(sheet_name, row_idx, field): new_value}` to
   `{(sheet_name, row_uid, field): new_value}` to reflect the new key schema.
   The runtime type is still `dict`; no structural change needed.

---

**File**: `webapp/models/requests.py`

**Specific Changes**:
1. **`CellEditRequest`**: Replace `row_index: int` with `row_uid: str`.
2. **`DeleteRowRequest`**: Replace `row_indices: List[int]` with `row_uids: List[str]`.

---

**File**: `webapp/models/responses.py`

**Specific Changes**:
1. **`CellEditResponse`**: Replace `row_index: int` with `row_uid: str`.

---

**File**: `webapp/services/edit_service.py`

**Function**: `edit_cell`

**Specific Changes**:
1. **Look up row by `_row_uid`**: Replace the `req.row_index` bounds check and
   `sheet.rows[req.row_index]` access with a linear scan:
   ```python
   row_idx = next(
       (i for i, r in enumerate(sheet.rows) if r.get("_row_uid") == req.row_uid),
       None,
   )
   if row_idx is None:
       raise HTTPException(status_code=404, detail=f"Row '{req.row_uid}' not found.")
   ```
2. **Store edit keyed by `row_uid`**: Change
   `record.edits[(sheet_name, req.row_index, req.field_name)]` to
   `record.edits[(sheet_name, req.row_uid, req.field_name)]`.
3. **Return `row_uid` in response**: Change `CellEditResponse(row_index=..., ...)` to
   `CellEditResponse(row_uid=req.row_uid, ...)`.

**Function**: `delete_rows`

**Specific Changes**:
1. **Accept `row_uids`**: Replace `req.row_indices` with `req.row_uids`.
2. **Look up rows by `_row_uid`**: Build the set of indices to delete by scanning
   `sheet.rows` for matching `_row_uid` values, then delete in reverse index order
   as before.

---

**File**: `webapp/services/normalization_service.py`

**Function**: `normalize` (edit replay block)

**Specific Changes**:
1. **Match edits by `row_uid`**: Replace the current replay logic:
   ```python
   # OLD
   if 0 <= edit_row < len(sheet_obj.rows):
       if edit_field in sheet_obj.rows[edit_row]:
           sheet_obj.rows[edit_row][edit_field] = edit_value
   ```
   with:
   ```python
   # NEW  (edit_row is now a row_uid string)
   for row in sheet_obj.rows:
       if row.get("_row_uid") == edit_row and edit_field in row:
           row[edit_field] = edit_value
           break
   ```
   The variable name `edit_row` in the loop unpacking `(edit_sheet, edit_row, edit_field)`
   now holds a `row_uid` string rather than an integer index.

---

**File**: `webapp/services/workbook_service.py`

**Function**: `get_sheet_data`

**Specific Changes**:
1. **Do NOT strip `_row_uid` and `_excel_row_number` from `clean_rows`**: The
   current comprehension strips all underscore-prefixed keys:
   ```python
   clean_rows = [{k: v for k, v in row.items() if not k.startswith("_")} ...]
   ```
   Change the exclusion predicate to allow `_row_uid` and `_excel_row_number`
   through while still stripping all other internal keys:
   ```python
   _IDENTITY_KEYS = {"_row_uid", "_excel_row_number"}
   clean_rows = [
       {k: v for k, v in row.items()
        if not k.startswith("_") or k in _IDENTITY_KEYS}
       for row in sheet.rows
   ]
   ```
   These two fields will be present in the JSON response so the frontend can
   store them on `<tr>` elements.  They are underscore-prefixed so the existing
   `display_columns` logic (which only includes non-underscore keys) will
   naturally exclude them from the visible column list.

---

**File**: `webapp/services/export_service.py`

**Function**: `visible_rows`

**Specific Changes**:
1. **Preserve identity fields through filtering**: The current comprehension that
   strips `_normalization*` keys must also preserve `_row_uid` and
   `_excel_row_number`:
   ```python
   rows = [
       {k: v for k, v in row.items()
        if not k.startswith("_normalization")}
       for row in sheet_dataset.rows
   ]
   ```
   This already preserves `_row_uid` and `_excel_row_number` because they do not
   start with `_normalization`.  No change needed here — but verify that
   `apply_derived_columns` does not strip them.

2. **Export loop — no change needed**: The export loop reads values via
   `EXPORT_MAPPING` which only references named business fields.  `_row_uid` and
   `_excel_row_number` are not in `EXPORT_MAPPING` and are therefore never written
   to output cells.  The loop is already correct once `visible_rows` preserves the
   identity fields.

---

**File**: `webapp/static/app.js`

**Function**: `renderGrid`

**Specific Changes**:
1. **Store `_row_uid` on each `<tr>`**: Replace:
   ```js
   const rowIndex = sheetData.rows.indexOf(row);
   const tr = document.createElement('tr');
   tr.dataset.rowIndex = rowIndex;
   ```
   with:
   ```js
   const tr = document.createElement('tr');
   tr.dataset.rowUid = row._row_uid;
   ```
   Remove the `rowIndex` variable entirely from this scope.

2. **Pass `row_uid` to delete button**: Replace
   `delBtn.addEventListener('click', () => deleteSingleRow(rowIndex))` with
   `delBtn.addEventListener('click', () => deleteSingleRow(row._row_uid))`.

3. **Pass `row_uid` to cell click**: Replace
   `td.addEventListener('click', () => makeEditable(td, rowIndex, col))` with
   `td.addEventListener('click', () => makeEditable(td, row._row_uid, col))`.

4. **Fix checkbox selection**: Replace `state.selectedRows.add(rowIndex)` /
   `state.selectedRows.delete(rowIndex)` with `state.selectedRows.add(row._row_uid)` /
   `state.selectedRows.delete(row._row_uid)`.  `state.selectedRows` changes from a
   `Set<number>` to a `Set<string>`.

**Function**: `makeEditable`

**Specific Changes**:
1. **Rename parameter**: `rowIndex` → `rowUid` (string).
2. **Send `row_uid` in API call**: Replace `{ row_index: rowIndex, ... }` with
   `{ row_uid: rowUid, ... }`.
3. **Update in-memory row**: Replace `state.sheetData.rows[rowIndex]` with a lookup:
   ```js
   const row = state.sheetData.rows.find(r => r._row_uid === rowUid);
   if (row) row[fieldName] = newValue;
   ```

**Function**: `deleteSingleRow` / `deleteSelectedRows` / `_deleteRows`

**Specific Changes**:
1. **`deleteSingleRow(rowUid)`**: Parameter is now a `row_uid` string.
2. **`_deleteRows(uids)`**: Send `{ row_uids: uids }` instead of
   `{ row_indices: indices }`.
3. **In-memory splice after delete**: Replace index-based splice with uid-based
   filter:
   ```js
   const uidSet = new Set(uids);
   state.sheetData.rows = state.sheetData.rows.filter(r => !uidSet.has(r._row_uid));
   ```
4. **`toggleSelectAll`**: Uses `row._row_uid` instead of `tr.dataset.rowIndex`.

---

## Testing Strategy

### Validation Approach

The testing strategy follows a two-phase approach: first, surface counterexamples
that demonstrate the bug on unfixed code, then verify the fix works correctly and
preserves existing behavior.

### Exploratory Bug Condition Checking

**Goal**: Surface counterexamples that demonstrate the bug BEFORE implementing the
fix.  Confirm or refute the root cause analysis.  If we refute, we will need to
re-hypothesize.

**Test Plan**: Write tests that construct a `SheetDataset` with a leading empty row
(or numeric helper row) followed by data rows, apply the same filters that
`WorkbookService.get_sheet_data` applies, record an edit against the filtered-view
index, and then assert that the edit lands on the correct source row.  Run these
tests on the UNFIXED code to observe failures.

**Test Cases**:
1. **Helper-row shift test**: Sheet has rows [helper, "Yotam", "Rachel"]; after
   filter, "Rachel" is at filtered index 1.  Edit `(sheet, 1, "first_name_corrected")`
   to "Racheli".  Assert that the backing row for "Rachel" (not "Yotam") is updated.
   (Will fail on unfixed code — edit hits "Yotam".)

2. **Empty-row shift test**: Sheet has rows [empty, "David", "Sara"]; after
   empty-row filter, "David" is at filtered index 0.  Edit `(sheet, 0, "first_name_corrected")`
   to "Davidi".  Assert that "David"'s row is updated, not the empty row.
   (Will fail on unfixed code — edit hits wrong row.)

3. **Re-normalization replay test**: Record an edit against row at index 2, then
   delete row at index 0, then re-normalize.  Assert that the edit is replayed on
   the correct source row (now at index 1).
   (Will fail on unfixed code — replay hits wrong row.)

4. **Export position test**: Sheet has rows [helper, "Yotam", "Rachel"]; edit
   "Rachel" to "Racheli"; export.  Assert that the exported row for Excel row 4
   contains "Racheli" and Excel row 3 contains the original "Yotam" value.
   (Will fail on unfixed code — export writes "Racheli" to the wrong output row.)

**Expected Counterexamples**:
- Edit values appear in the wrong row of the exported workbook.
- Re-normalization replay overwrites the wrong row.
- Possible causes: index shift from filter, index-keyed edit store, sequential
  export output position.

### Fix Checking

**Goal**: Verify that for all inputs where the bug condition holds, the fixed
pipeline produces the expected behavior.

**Pseudocode:**
```
FOR ALL pipeline_state WHERE isBugCondition(pipeline_state) DO
  result := export(pipeline_state_fixed)
  ASSERT result[_excel_row_number = edited_row._excel_row_number] contains edited_value
  ASSERT no_other_row_modified(result)
END FOR
```

### Preservation Checking

**Goal**: Verify that for all inputs where the bug condition does NOT hold, the
fixed pipeline produces the same result as the original pipeline.

**Pseudocode:**
```
FOR ALL pipeline_state WHERE NOT isBugCondition(pipeline_state) DO
  ASSERT export_original(pipeline_state) = export_fixed(pipeline_state)
END FOR
```

**Testing Approach**: Property-based testing is recommended for preservation
checking because:
- It generates many test cases automatically across the input domain.
- It catches edge cases that manual unit tests might miss.
- It provides strong guarantees that behavior is unchanged for all non-buggy inputs.

**Test Plan**: Observe behavior on UNFIXED code first for workbooks with no skipped
rows, then write property-based tests capturing that behavior.

**Test Cases**:
1. **No-skip preservation**: Generate random `SheetDataset` rows with no empty rows
   and no numeric helper row.  Assert that the fixed pipeline produces identical
   normalized field values and identical export output.

2. **Edit-then-export preservation (no skips)**: Generate a sheet, apply a random
   edit, export.  Assert that the exported value matches the edit and all other
   rows are unchanged — same as before the fix.

3. **Multi-sheet preservation**: Generate a workbook with multiple sheets, none
   with skipped rows.  Assert that each sheet's export output is identical before
   and after the fix.

4. **Normalization output preservation**: For any sheet, assert that
   `_corrected` and `_status` field values produced by the pipeline are identical
   before and after the fix (identity fields are internal and do not affect
   normalization output).

### Unit Tests

- Test that `extract_sheet_to_json` assigns `_row_uid`, `_excel_row_number`,
  `_source_data_index`, and `_source_sheet` to every extracted row.
- Test that `_row_uid` equals `"{sheet_name}:{row_num}"` for each row.
- Test that `normalize_dataset` preserves all four identity fields unchanged.
- Test that `edit_cell` with a valid `row_uid` updates the correct row and stores
  the edit keyed by `row_uid`.
- Test that `edit_cell` with an unknown `row_uid` returns HTTP 404.
- Test that `delete_rows` with valid `row_uids` removes the correct rows.
- Test that `NormalizationService.normalize` replays edits by `row_uid` after
  re-normalization, even when the array order has changed.
- Test that `WorkbookService.get_sheet_data` includes `_row_uid` and
  `_excel_row_number` in the row dicts of the response.
- Test that `visible_rows` preserves `_row_uid` and `_excel_row_number` through
  all filtering steps.

### Property-Based Tests

- Generate random `SheetDataset` instances (varying row counts, field values,
  presence of empty rows and helper rows) and verify that every row in the
  extracted dataset has a unique `_row_uid`.
- Generate random edit sequences on sheets with skipped rows and verify that each
  edit lands on the row with the matching `_row_uid` in the export output.
- Generate random non-buggy sheets (no skips) and verify that the fixed pipeline's
  export output is byte-for-byte identical to the original pipeline's output for
  all business fields.
- Generate random `row_uid` strings and verify that `edit_cell` with an unknown
  uid always returns 404 (never silently corrupts another row).

### Integration Tests

- Upload a workbook with a numeric helper row, normalize, edit a cell in the UI
  (simulated via API), export, and assert the correct Excel row contains the
  edited value.
- Upload a workbook, delete a row via the API, re-normalize, export, and assert
  that no other row's data was corrupted.
- Upload a multi-sheet workbook, edit cells on two different sheets, export, and
  assert that both edits appear in the correct rows of the correct sheets.
- Verify that the UI grid does not display `_row_uid` or `_excel_row_number` as
  visible columns.
