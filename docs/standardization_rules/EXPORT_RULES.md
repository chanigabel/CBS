# Export Rules

## 1. Purpose

Define how normalized workbook data is converted to the final Excel export
workbook.

## 2. Scope

Applies to:

- active web export service and writer
- compatibility `ExportEngine`
- export row filtering and derived columns

Implemented by:

- `webapp/services/export_service.py`
- `webapp/services/export_writer.py`
- `webapp/services/export_rows.py`
- `webapp/services/export_schema.py`
- `webapp/services/export_validation.py`
- `src/excel_standardization/export/export_engine.py`

## 3. Source Fields

Export reads normalized `SheetDataset.rows`, display-visible source fields,
corrected fields, session metadata, and sheet metadata.

## 4. Corrected Fields

Active web export maps final columns from corrected fields:

- `ShemPrati` -> `first_name_corrected`
- `ShemMishpaha` -> `last_name_corrected`
- `ShemHaAv` -> `father_name_corrected`
- `MisparZehut` -> `id_number_corrected`
- `Darkon` -> `passport_corrected`
- `Min` -> `gender_corrected`
- birth/entry export components -> corrected date component fields

## 5. Status Fields

The final export workbook does not include standardization status fields or
`_validation_status` in the active schema.

## 6. Original-Value Immutability Rules

Export must not mutate original source files. Web export may mutate in-memory row
dictionaries by injecting `MosadID` and `SugMosad` before writing.

## 7. Corrected-Field Contract

The active web export schema expects corrected fields to exist after
standardization. If a mapped corrected value is missing or empty, the cell is
left empty; active web export does not use original fallback through
`EXPORT_MAPPING`.

Compatibility `ExportEngine._map_row_to_export_fields` also maps standardized
columns from corrected fields only. If a corrected standardized value is missing
or empty, the exported cell remains empty. Mosad/apartment metadata may still
use explicit metadata/source fallbacks.

## 8. Parsing / Cleanup / Normalization Rules

Export performs no engine-level cleanup. It only:

- canonicalizes sheet names
- filters visible rows
- injects derived columns and institution metadata
- maps JSON keys to export headers
- writes workbook sheets and headers

## 9. Validation Rules

Before writing web export rows:

- rows without any visible original values are removed
- the first row is removed if all non-empty original field values are numeric-like
  helper values

Compatibility `ExportEngine` exports only rows where any one of key personal
fields is non-empty.

## 10. Recovery Rules

If session workbook data is missing, web export attempts to re-extract the
working copy. If export writing fails, the service records a processing-report
error and raises HTTP 500 while preserving session data.

## 11. Ambiguity Rules

Sheet-name canonicalization is keyword-based. Unknown sheet names are exported
under their source name using the default header schema in the web writer.

## 12. Invalid-Value Behavior

Export does not revalidate invalid values. It writes whatever mapped corrected
fields contain, leaving missing/empty values blank.

## 13. Export Behavior

Active web export:

- creates one output sheet per dataset sheet
- canonicalizes sheet titles to `DayarimYahidim`, `MeshkeyBayt`, or
  `AnasheyTzevet` when recognized
- writes headers in row 1
- sets worksheet right-to-left display
- writes data starting at row 2
- applies scoped or default `SugMosad`
- applies session `MosadID` when available
- returns row counts by sheet for the processing report

Export filename:

- If session has both MosadID and Mosad name, filename is
  `<MosadID> <PascalCaseMosadName>.xlsx`.
- Otherwise filename is `<original_stem>_standardized_<timestamp>.xlsx`.

Compatibility `ExportEngine`:

- always creates fixed sheets for the three known source sheet specs
- has no-Dira and with-Dira schema variants
- can append extra source sheets
- has worksheet-based export for augmented workbooks and JSON-based export for
  normalized datasets

## 14. UI/Grid Behavior

Export visible rows use the same `visible_rows` helper as the UI for row filtering
and derived serial/MosadID columns.

## 15. API Behavior

The web export endpoint calls `ExportService.export(session_id)`, writes the file
under the configured output directory, updates the processing report, and returns
the workbook for download.

## 16. Examples

| Export Header | Active Web Source |
|---|---|
| `MosadID` | session/sheet row `MosadID` |
| `SugMosad` | scoped config or active session Mosad type |
| `ShemPrati` | `first_name_corrected` |
| `Darkon` | `passport_corrected` |
| `ShnatLida` | `birth_year_corrected` |
| `YomKnisa` | `entry_day_corrected` |

## 17. Current Known Limitations

- Final export does not include status columns.
- Active web export and compatibility `ExportEngine` share corrected-only
  behavior for standardized output columns.
- Unknown sheets use default web headers, while compatibility export appends
  extra sheets differently.

## 18. Open Questions Requiring Approval

- Should export include a validation/status report sheet?
- Should Mosad/apartment metadata fallback rules be narrowed further?
- Should unknown sheets be exported with source columns or skipped?

## 19. Tests That Should Cover The Behavior

- `tests/test_export_engine_dataset.py`
- `tests/webapp/test_export_service.py`
- `tests/webapp/test_api_export.py`
- `tests/test_date_engine_corrected_flow.py`

## 20. Final Principles

Export should be a projection of standardized data, not a second
standardization engine. Corrected-field mapping must remain explicit.
