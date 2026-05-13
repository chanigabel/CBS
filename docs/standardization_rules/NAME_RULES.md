# Name Rules

## 1. Purpose

Normalize personal name fields while preserving original input values and
writing deterministic corrected fields.

## 2. Scope

Applies to:

- `first_name`
- `last_name`
- `father_name`

Implemented by:

- `src/excel_standardization/engines/name_engine.py`
- `src/excel_standardization/engines/text_processor.py`
- `src/excel_standardization/processing/name_standardization.py`
- `src/excel_standardization/processing/standardization_pipeline.py`

## 3. Source Fields

- `first_name`
- `last_name`
- `father_name`

Fields are processed only if the key exists in the row.

## 4. Corrected Fields

- `first_name_corrected`
- `last_name_corrected`
- `father_name_corrected`

## 5. Status Fields

Approved rule: no dedicated name status field is currently produced.

Current behavior: if name standardization raises an exception, the pipeline
records failed fields in `_standardization_failures` and falls back by copying
original values into missing corrected fields.

## 6. Original-Value Immutability Rules

Approved rule: original name fields must not be modified by standardization.
All normalized output goes into `*_corrected`.

## 7. Corrected-Field Contract

- If the original value is `None` or an empty string, the corrected field is set
  to that same value.
- Otherwise, the corrected field receives the cleaned name output.
- On exception, the corrected field falls back to the original value and the
  field is listed in `_standardization_failures`.

## 8. Parsing / Cleanup / Normalization Rules

Name cleanup delegates to `TextProcessor.clean_name`.

Approved rule:

- Convert input safely to string.
- Remove zero-width/invisible Unicode characters.
- Remove supported diacritics.
- Translate Arabic-Indic digits before filtering.
- Detect dominant language by counting Hebrew and English letters.
- Hebrew wins ties.
- Keep only dominant-language letters, spaces, and separators converted to spaces.
- Convert hyphen-like characters, parentheses, and backslash to spaces.
- Drop digits, symbols, and non-dominant-language letters.
- Collapse whitespace.
- Remove configured Hebrew/English unwanted title tokens after character filtering.
- Remove parenthesized acronym groups only when the parenthesized text contains
  a quote/acronym character.

## 9. Last-Name Removal Rules

Approved rule: first-name and father-name cleanup can remove an embedded last
name using a two-stage process.

Dataset-level pattern detection:

- Samples up to five rows.
- Requires at least three rows where the last name appears in the target field.
- If at least three matches are first token matches, pattern is `REMOVE_FIRST`.
- If at least three matches are last token matches, pattern is `REMOVE_LAST`.
- Otherwise pattern is `NONE`.

Direct row-local removal:

- Direct last-name removal is always applied independently of dataset-level
  pattern detection.
- If the current row’s `last_name` appears inside `first_name` or `father_name`,
  the embedded last-name substring is removed.
- This applies to:
  - single-word last names
  - multi-word last names
- Removal uses normalized whole-word space-padded matching.

Examples:

- `first_name="Jacob Cohen"`, `last_name="Cohen"` -> `Jacob`
- `father_name="Abraham Ben David"`, `last_name="Ben David"` -> `Abraham`

First name:

- Single-word first names are never modified by positional fallback.
- Stage A performs direct row-local last-name removal.
- If Stage A changes the value, Stage B does not run.
- If Stage A does not change the value and the pattern is not `NONE`, Stage B
  removes the first or last token according to the detected pattern.

Father name:

- Single-word father names are never modified by positional fallback.
- Stage A performs direct row-local last-name removal.
- If Stage A changes the value, Stage B does not run.
- If Stage A does not change the value and the pattern is not `NONE`, Stage B
  removes the first or last token according to the detected pattern.
  
## 10. Validation Rules

Name engine validation is cleanup-only. Required-name validation is handled by
`InstitutionReportValidator`:

- `first_name_corrected` or fallback `first_name` must be non-empty for
  institution reports.
- `last_name_corrected` or fallback `last_name` must be non-empty.

## 11. Recovery Rules

On engine exception:

- missing corrected fields are filled with their original values
- failed fields are listed in `_standardization_failures`
- processing continues for the row and dataset

## 12. Ambiguity Rules

Current behavior: last-name removal ambiguity is resolved by dataset-level
pattern detection. If the sample is insufficient, no positional removal is
performed.

Needs approval: there is no user-visible status for ambiguous name cleanup.

## 13. Invalid-Value Behavior

Current behavior: values that clean down to an empty string remain empty in the
corrected field. No name-specific status is written.

## 14. Export Behavior

The active web export maps:

- `ShemPrati` from `first_name_corrected`
- `ShemMishpaha` from `last_name_corrected`
- `ShemHaAv` from `father_name_corrected`

The compatibility `ExportEngine` uses corrected fields first and may fall back
to original fields in `_map_row_to_export_fields`.

Potential issue: fallback-to-original in compatibility export should be reviewed
if strict corrected-only export is required outside date fields.

## 15. UI/Grid Behavior

The grid shows original name fields in source order and places each existing
`*_corrected` field immediately after its original field.

## 16. API Behavior

The `/standardize` flow builds `NameEngine(TextProcessor())` and runs name
standardization for loaded sheets. Manual edits are replayed after
standardization when recorded in the session.

## 17. Examples

| Source | Corrected | Status |
|---|---|---|
| `Smith-Jones` | `Smith Jones` | none |
| Hebrew name plus digits | Hebrew letters only | none |
| `first_name="Cohen Jacob"`, `last_name="Cohen"`, pattern `REMOVE_FIRST` | `Jacob` | none |
| `father_name="Abraham Cohen"`, `last_name="Cohen"`, pattern `REMOVE_LAST` | `Abraham` | none |

## 18. Current Known Limitations

- No name-specific status field.
- Pattern detection depends on early rows and may not represent the whole sheet.
- Last-name substring removal is not a linguistic parser; it uses deterministic
  string/token rules.

## 19. Open Questions Requiring Approval

- Should ambiguous name cleanup write a visible status?
- Should export ever fall back to original name fields when corrected fields are
  absent or empty?
- Should first-name and father-name pattern detection sample more than five rows?

## 20. Tests That Should Cover The Behavior

- `tests/test_name_engine.py`
- `tests/test_normalization_pipeline.py`
- `tests/test_institution_report_validator.py`
- web UI tests that verify corrected columns appear beside originals

## 21. Final Principles

Names are cleaned deterministically, not inferred. Original names remain
immutable. If the code cannot safely identify a cleanup pattern, it should leave
the value unchanged or document the behavior as needing approval.
