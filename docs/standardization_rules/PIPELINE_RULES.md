# Pipeline Rules

## 1. Purpose

Document the orchestration contract for extraction, standardization, validation,
UI presentation, API behavior, and export.

## 2. Scope

Applies to:

- `StandardizationPipeline`
- web `StandardizationService`
- workbook UI data service
- processing report stages
- post-standardization validation

Implemented by:

- `src/excel_standardization/processing/standardization_pipeline.py`
- `src/excel_standardization/processing/*_standardization.py`
- `src/excel_standardization/workbook_json_flow.py`
- `webapp/services/standardization_service.py`
- `webapp/services/workbook_service.py`
- `webapp/services/export_service.py`

## 3. Source Fields

The pipeline accepts `SheetDataset.rows` produced by extraction. Source field
names are normalized by the extraction layer and stored in `SheetDataset.field_names`.

## 4. Corrected Fields

Corrected fields use the `_corrected` suffix. Active engines write:

- names: `first_name_corrected`, `last_name_corrected`, `father_name_corrected`
- gender: `gender_corrected`
- identifiers: `id_number_corrected`, `passport_corrected`
- dates: corrected year/month/day component fields for birth and entry dates

## 5. Status Fields

Engine/status fields:

- `gender_status`
- `identifier_status`
- `birth_date_status`
- `entry_date_status`
- `_standardization_failures`
- `_validation_status`
- `_validation_ok`

Internal helper fields beginning with `_` are implementation details unless
explicitly kept by UI behavior.

## 6. Original-Value Immutability Rules

Approved rule: standardization copies the input row before applying engines.
Original source fields should remain unchanged; all standardized output is added
as corrected/status fields.

## 7. Corrected-Field Contract

Each engine only writes fields for source fields/groups that exist or are
created by current pairwise logic. Missing engines or disabled flags skip the
corresponding standardization.

## 8. Parsing / Cleanup / Normalization Rules

Pipeline order:

1. Name standardization
2. Gender standardization
3. Date standardization
4. Identifier standardization
5. Dataset-level birth-year majority correction
6. Sheet-level institution validation for known sheets
7. Metadata/statistics update

Current behavior: validation runs after corrected fields exist, so validators can
prefer corrected values.

## 9. Validation Rules

The pipeline collects standardization exceptions into `_standardization_failures`
and metadata statistics. Institution validation adds row-level validation
statuses after engine standardization.

## 10. Recovery Rules

Engine helpers catch exceptions locally where implemented and allow other engines
or rows to continue. The web service catches sheet-level standardization failure,
adds a processing warning, and continues if at least one sheet succeeds.

## 11. Ambiguity Rules

Current behavior:

- Name patterns and date format patterns are detected once per dataset from row
  samples.
- Date format defaults to DDMM unless MMDD evidence exceeds DDMM evidence.
- Validation skips rules requiring unavailable external reference data.

## 12. Invalid-Value Behavior

Invalid values should not overwrite original fields. Engines write corrected
empty values and visible status fields according to each engine's contract.

## 13. Export Behavior

Export reads normalized datasets from session state. Active web export expects
the standardization service to have produced corrected fields; it does not rerun
standardization.

## 14. UI/Grid Behavior

Workbook UI data:

- assigns stable `_row_uid` values
- strips underscore-prefixed internal keys except `_row_uid` and `_validation_status`
- drops empty rows using original source fields only
- hides the first numeric helper row
- places corrected fields beside originals
- groups date source fields, corrected components, and date status together
- appends `_validation_status` when present
- injects serial, MosadID, and SugMosad display columns

## 15. API Behavior

Standardization API:

- lazily extracts workbook/sheet data if not loaded
- runs all active engines
- merges normalized sheets back into session state
- runs workbook-level institution validation
- replays recorded manual edits after standardization
- updates session status to `standardized`

Export API:

- writes the current session dataset to a new workbook
- updates processing report export details

## 16. Examples

| Stage | Example Output |
|---|---|
| name engine | `first_name_corrected` |
| gender engine | `gender_corrected`, optionally `gender_status` |
| date engine | corrected date components and date status |
| identifier engine | `id_number_corrected`, `passport_corrected`, `identifier_status` |
| validation | `_validation_status`, `_validation_ok` |
| UI | original/corrected/status columns in display order |
| export | fixed schema workbook |

## 17. Current Known Limitations

- Several statuses are display strings only, not status codes.
- UI and export share visible-row filtering but final export does not include
  status fields.
- Manual edit replay after standardization may overwrite standardized fields if
  the edited field is the same key.

## 18. Open Questions Requiring Approval

- Should all engines emit stable status codes?
- Should manual edits apply before or after standardization depending on field
  type?
- Should status fields be included in a secondary export/report artifact?

## 19. Tests That Should Cover The Behavior

- `tests/test_normalization_pipeline.py`
- `tests/webapp/test_normalization_service.py`
- `tests/webapp/test_workbook_service.py`
- `tests/webapp/test_integration_webapp.py`
- engine-specific test files

## 20. Final Principles

The pipeline coordinates engines; it should not hide engine-specific uncertainty.
Original fields remain stable, corrected fields are explicit, and status fields
carry visible explanations.
