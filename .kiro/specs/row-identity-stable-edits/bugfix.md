# Bugfix Requirements Document

## Introduction

A critical data corruption bug exists in the Excel normalization pipeline. When one or more rows are removed, skipped, hidden, or filtered out at any stage of the pipeline (extraction, display, editing, or export), user edits are written back to the wrong source rows in the exported Excel file.

**Concrete example:** Original Excel row 4 contains "Yotam", row 5 contains "Rachel". If a row before them is skipped (e.g. a numeric helper row or an empty row), the in-memory array shifts so that "Rachel" appears at array index 0 and "Yotam" at array index 1. When the user edits "Rachel" to "Racheli", the edit is stored against array index 0. On export, the system writes "Racheli" into the wrong source row — corrupting data that was never touched.

The root cause is that the entire pipeline — from extraction through normalization, display, editing, and export — uses array index, visible row index, or filtered-dataset position as the row identity. None of these are stable when rows are removed or reordered. The fix must introduce a stable, source-anchored row identity at the earliest extraction stage and preserve it through every subsequent stage.

This is a critical infrastructure bug, not a normalization-engine issue.

---

## Bug Analysis

### Current Behavior (Defect)

1.1 WHEN a row is extracted from Excel and one or more preceding rows are later skipped (empty-row filter, numeric helper-row filter, or user deletion), THEN the system uses the remaining array index to identify rows, causing all subsequent rows to shift their identity by the number of skipped rows.

1.2 WHEN the user edits a cell in the UI, THEN the system stores the edit keyed by the row's current array index in `sheet.rows`, which is not stable across row removals or re-normalizations.

1.3 WHEN the export pipeline calls `visible_rows()` to filter and reorder rows, THEN the system iterates the resulting list by position and writes each row to the next sequential output row, with no reference to the original Excel row number.

1.4 WHEN normalization is re-run after manual edits, THEN the system replays edits by `(sheet_name, row_idx, field_name)` where `row_idx` is the array index at the time of the original edit, which may no longer point to the same source row if rows were deleted or reordered in between.

1.5 WHEN the UI applies column filters or sorting before the user makes an edit, THEN the `rowIndex` sent to the API is derived from `sheetData.rows.indexOf(row)` on the filtered/sorted view, which may differ from the row's true position in the backing dataset.

### Expected Behavior (Correct)

2.1 WHEN a row is extracted from Excel, THEN the system SHALL assign each row a stable identity comprising `_source_sheet` (sheet name), `_excel_row_number` (1-based Excel row number), `_source_data_index` (0-based position in the raw extracted array before any filtering), and `_row_uid` (string `"{sheetName}:{excelRowNumber}"`), and these fields SHALL be stored on the row dict at extraction time.

2.2 WHEN the user edits a cell in the UI, THEN the system SHALL store the edit keyed by `_row_uid` (not by array index), so the edit is permanently bound to the exact source row regardless of any subsequent row removal, reordering, or re-normalization.

2.3 WHEN the export pipeline writes data rows to the output workbook, THEN the system SHALL look up each row's `_excel_row_number` (or `_row_uid`) and write the corrected values to that exact row in the source workbook, never by sequential output position.

2.4 WHEN normalization is re-run after manual edits, THEN the system SHALL replay edits by matching `_row_uid`, so edits survive re-normalization even if the array order of rows has changed.

2.5 WHEN the UI applies column filters, sorting, or any other view transformation before the user makes an edit, THEN the system SHALL send `_row_uid` (not `rowIndex`) to the API, and the API SHALL locate the target row by `_row_uid` in the backing dataset.

2.6 WHEN any row removal occurs (empty-row filter, numeric helper-row filter, user deletion, or hidden-row skip), THEN the system SHALL NOT alter the `_row_uid` or `_excel_row_number` of any other row.

### Unchanged Behavior (Regression Prevention)

3.1 WHEN all rows are present with no skips, filters, or deletions, THEN the system SHALL CONTINUE TO extract, normalize, display, and export all rows correctly with the same field values as before.

3.2 WHEN the user edits a corrected field on a row that has no preceding skipped rows, THEN the system SHALL CONTINUE TO write the edited value to the correct output row on export.

3.3 WHEN normalization is run on a sheet, THEN the system SHALL CONTINUE TO produce the same corrected field values (`_corrected`, `_status`) for each row as before.

3.4 WHEN the export service builds the output workbook, THEN the system SHALL CONTINUE TO apply the canonical sheet-name mapping, the fixed column schema, and the right-to-left sheet direction.

3.5 WHEN the UI displays sheet data, THEN the system SHALL CONTINUE TO hide internal metadata keys (underscore-prefixed fields) from the rendered grid.

3.6 WHEN the numeric helper row or empty rows are filtered out for display, THEN the system SHALL CONTINUE TO hide those rows from the UI grid and from the export output, while preserving the stable identity of all other rows.

3.7 WHEN multiple sheets are present in a workbook, THEN the system SHALL CONTINUE TO process each sheet independently with correct row identity scoped to its own sheet name.

---

## Bug Condition (Pseudocode)

```pascal
FUNCTION isBugCondition(pipeline_state)
  INPUT: pipeline_state describing the rows in memory and the edit store
  OUTPUT: boolean

  // The bug is triggered whenever the array index of a row in sheet.rows
  // differs from the index that was used when the edit was recorded,
  // OR when the export iterates rows by position rather than by source row number.

  RETURN (
    EXISTS row IN sheet.rows WHERE
      row does NOT have a stable _row_uid field
    OR
    EXISTS edit IN record.edits WHERE
      edit is keyed by (sheet_name, array_index, field) AND
      array_index no longer points to the same source row
    OR
    export writes row at output_position WITHOUT referencing _excel_row_number
  )
END FUNCTION
```

**Fix Checking Property:**
```pascal
// FOR ALL inputs where one or more rows are skipped before the edited row:
FOR ALL pipeline_state WHERE isBugCondition(pipeline_state) DO
  result ← export(pipeline_state)
  ASSERT result[source_row_4] = "Yotam"          // untouched row unchanged
  ASSERT result[source_row_5] = "Racheli"         // edited row updated correctly
  ASSERT no_other_row_modified(result)
END FOR
```

**Preservation Property:**
```pascal
// FOR ALL inputs where no rows are skipped (non-buggy path):
FOR ALL pipeline_state WHERE NOT isBugCondition(pipeline_state) DO
  ASSERT export_before_fix(pipeline_state) = export_after_fix(pipeline_state)
END FOR
```
