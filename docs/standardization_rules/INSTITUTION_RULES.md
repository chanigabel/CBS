# Institution Rules

## 1. Purpose

Document Mosad-related behavior and institution-report validation applied after
standardization.

## 2. Scope

Applies to recognized institution report sheets:

- `DayarimYahidim`
- `MeshkeyBayt`
- `AnasheyTzevet`

Implemented by:

- `src/excel_standardization/validation/institution_report_validator.py`
- `src/excel_standardization/services/sheet_name_resolver.py`
- `webapp/services/mosad_id_scanner.py`
- `webapp/services/derived_columns.py`
- `webapp/services/export_rows.py`
- `webapp/services/export_writer.py`
- `webapp/services/standardization_service.py`
- `webapp/services/workbook_service.py`

## 3. Source Fields

Validation reads original and corrected row fields, plus metadata/session values:

- `MosadID` or `mosad_id`
- `SugMosad` or `sug_mosad`
- `MisparDiraBeMosad`
- name, gender, identifier, birth date, and entry date fields
- sheet metadata `MosadID`
- session `mosad_id`
- session `mosad_types`

## 4. Corrected Fields

Institution validation does not create business corrected fields. It reads
`*_corrected` values when available.

## 5. Status Fields

- `_validation_status`: pipe-separated Hebrew validation messages.
- `_validation_ok`: boolean row validity based on error severity.

## 6. Original-Value Immutability Rules

Approved rule: validation must not overwrite original source fields or corrected
standardization fields. It mutates rows only by adding/updating internal
validation status keys.

## 7. Corrected-Field Contract

Validators prefer corrected values when present and non-empty, with fallback to
original values for name and date component checks. Identifier duplicate checks
use non-empty `id_number_corrected` only.

## 8. Parsing / Cleanup / Normalization Rules

Sheet names are resolved to canonical names through keyword matching:

- Hebrew Dayarim names map to `DayarimYahidim`.
- Hebrew Meshkey Bayt names map to `MeshkeyBayt`.
- Hebrew staff/family names map to `AnasheyTzevet`.
- Unknown names remain unchanged.

MosadID:

- Scanned from the workbook/sheet by `mosad_id_scanner`.
- Added to sheet metadata when found.
- Injected into UI/export rows from session or metadata.

SugMosad:

- Provided by session-level Mosad type values or scoped export configuration.
- Can be applied at workbook, sheet, or selected-row scope for export.

Serial number:

- UI/export visible rows add a serial column.
- Existing serial-like source columns are reused and blanks are filled.
- If no serial source column exists, synthetic `_serial` is added for display.

## 9. Validation Rules

Current implemented validation:

- `MosadID` is required but not checked for numeric format.
- `SugMosad` is required, numeric, and at least three digits.
- `MisparDiraBeMosad` is optional; if present outside `DayarimYahidim`, it must be numeric.
- First name is required after correction/fallback.
- Last name is required after correction/fallback.
- Non-empty corrected ID values must be unique within sheet.
- Non-empty corrected ID values duplicated across workbook sheets produce warning.
- Gender corrected value, when present, must be `1` or `2`.
- Birth year, month, and day are required and numeric.
- Birth year must be at least `1906` and not future.
- Birth month must be `1..12`.
- Birth day must be `1..31`.
- Entry year and month are required and numeric.
- Entry year must be less than or equal to validator `census_year`.
- Entry month must be `1..12`.
- Entry day is required only for `DayarimYahidim`.
- Entry day, when present, must be `1..31`.

## 10. Recovery Rules

If validation fails at the pipeline level, the standardization pipeline logs a
warning and continues. Workbook-level validation skipped by the web service is
reported as a processing-report warning.

## 11. Ambiguity Rules

Current behavior:

- Missing external reference dictionaries are not simulated.
- Duplicate cross-workbook ID findings are warnings, not errors.
- Empty corrected IDs are excluded from duplicate checks.

## 12. Invalid-Value Behavior

Invalid values write messages into `_validation_status`. Multiple messages are
pipe-separated. `_validation_ok` is false when any finding has severity `error`.

## 13. Export Behavior

Validation status is not mapped to the final export schema. Export injects
MosadID and SugMosad into rows before writing. `MisparDiraBeMosad` is exported
for Meshkey Bayt and Anashey Tzevet schemas.

## 14. UI/Grid Behavior

The UI keeps `_validation_status` visible and appends it to display columns when
present. `_validation_ok` remains internal and is stripped from displayed rows.

MosadID and SugMosad are injected into the grid from session/metadata and placed
near the front of display columns after the serial column.

## 15. API Behavior

The standardization service runs sheet-level validation during pipeline
normalization and workbook-level validation after merging normalized sheets back
into the session dataset.

## 16. Examples

| Case | Current Behavior |
|---|---|
| missing MosadID | `_validation_status` includes missing MosadID |
| `SugMosad="10"` | too-short SugMosad finding |
| duplicate corrected ID in same sheet | second occurrence gets duplicate error |
| corrected ID appears on two known sheets | warning on affected rows |
| missing entry day on Dayarim | validation error |
| missing entry day on Meshkey/Anashey | allowed |

## 17. Current Known Limitations

- External SugMosad dictionary membership is not implemented.
- Related-institution registry checks are not implemented.
- Minimum-entry-age by SugMosad is documented in code as unavailable.
- MosadID format is not validated beyond presence.

## 18. Open Questions Requiring Approval

- Should MosadID be numeric and length-validated?
- Should cross-sheet duplicate IDs be errors instead of warnings?
- Which external reference datasets should be integrated for SugMosad and
  minimum-entry-age checks?
- Should `_validation_status` be exported to a separate report workbook?

## 19. Tests That Should Cover The Behavior

- `tests/test_institution_report_validator.py`
- `tests/webapp/test_api_institution.py`
- `tests/webapp/test_export_service.py`
- `tests/webapp/test_workbook_service.py`

## 20. Final Principles

Institution validation is a post-standardization layer. It should report
findings without rewriting corrected standardization output.
