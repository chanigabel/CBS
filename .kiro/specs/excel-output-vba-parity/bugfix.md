# Bugfix Requirements Document

## Introduction

The Python normalization pipeline currently exports a **separate** output Excel file instead of augmenting the original workbook in-place. This diverges from the VBA implementation, which always modifies the original workbook by inserting corrected columns immediately to the right of each original column, then saves the same file.

The fix must restore full parity with the VBA output strategy: load the original workbook, insert new corrected columns adjacent to their originals, apply conditional highlighting, and save the modified workbook — never producing a separate file.

## Bug Analysis

### Current Behavior (Defect)

1.1 WHEN the pipeline processes an Excel workbook THEN the system writes normalized values to a **separate** output file instead of modifying the original workbook

1.2 WHEN a normalized field is produced THEN the system does NOT insert a new column immediately to the right of the original column; instead it writes to a pre-existing or separate sheet

1.3 WHEN the pipeline finishes THEN the system saves a new/separate Excel file, leaving the original workbook unmodified

1.4 WHEN a date field is processed THEN the system does NOT insert the four corrected sub-columns (שנה - מתוקן | חודש - מתוקן | יום - מתוקן | סטטוס תאריך) immediately after the original date columns

1.5 WHEN an identifier field is processed THEN the system does NOT insert the three corrected columns (ת.ז. - מתוקן | דרכון - מתוקן | סטטוס מזהה) immediately after the passport column

1.6 WHEN a corrected cell value equals Trim(OriginalInput) THEN the system incorrectly applies the pink highlight (RGB 255, 199, 206) to that cell

### Expected Behavior (Correct)

2.1 WHEN the pipeline processes an Excel workbook THEN the system SHALL load the original workbook and modify it in-place, never creating a separate output file

2.2 WHEN a normalized field is produced THEN the system SHALL insert a new column immediately to the RIGHT of the original column, with header equal to the original header text + " - מתוקן"

2.3 WHEN the pipeline finishes THEN the system SHALL save the modified original workbook (augmented, not replaced)

2.4 WHEN a date field is processed THEN the system SHALL insert exactly four corrected columns immediately after the original day column, with headers: "שנה - מתוקן", "חודש - מתוקן", "יום - מתוקן", "סטטוס תאריך"

2.5 WHEN an identifier field is processed THEN the system SHALL insert exactly three corrected columns immediately after the passport column, with headers: "ת.ז. - מתוקן", "דרכון - מתוקן", "סטטוס מזהה"

2.6 WHEN a corrected cell value differs from Trim(OriginalInput) THEN the system SHALL apply pink highlight RGB(255, 199, 206) to that corrected cell

2.7 WHEN a corrected cell value equals Trim(OriginalInput) THEN the system SHALL NOT apply any highlight to that corrected cell

2.8 WHEN inserting corrected columns THEN the system SHALL preserve ALL original columns exactly as-is (no overwriting, reordering, modifying, or deleting of original data)

### Unchanged Behavior (Regression Prevention)

3.1 WHEN a name field (שם פרטי, שם משפחה, שם האב) is processed THEN the system SHALL CONTINUE TO apply text normalization (trim, diacritics removal, language dominance, space collapsing) to produce the corrected value

3.2 WHEN a gender field is processed THEN the system SHALL CONTINUE TO normalize gender values to 1 (male) or 2 (female) using the existing pattern matching logic

3.3 WHEN a date field is processed THEN the system SHALL CONTINUE TO parse dates from split columns or main value using the existing DateEngine logic (8-digit, 6-digit, 4-digit, separated formats)

3.4 WHEN an identifier field is processed THEN the system SHALL CONTINUE TO validate Israeli IDs with checksum verification and clean passport values using the existing IdentifierEngine logic

3.5 WHEN a workbook contains multiple worksheets THEN the system SHALL CONTINUE TO process each worksheet sequentially in the order: names → gender → dates → identifiers

3.6 WHEN a date status cell has non-empty status containing "גיל מעל" THEN the system SHALL CONTINUE TO apply yellow background (RGB 255, 230, 150) and bold font

3.7 WHEN a date status cell has non-empty status not containing "גיל מעל" THEN the system SHALL CONTINUE TO apply pink background (RGB 255, 200, 200) and bold font

3.8 WHEN a header is not found in a worksheet THEN the system SHALL CONTINUE TO skip processing for that field without error
