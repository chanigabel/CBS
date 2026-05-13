# Identifier Rules

## 1. Purpose

Normalize Israeli ID and passport fields together, validate Israeli ID checksum,
and route passport-like values to the passport corrected field when current
rules require it.

## 2. Scope

Applies to:

- `id_number`
- `passport`

Implemented by:

- `src/excel_standardization/engines/identifier_engine.py`
- `src/excel_standardization/processing/identifier_standardization.py`
- `src/excel_standardization/validation/institution_report_validator.py`

## 3. Source Fields

- `id_number`
- `passport`

The identifier processor runs when either source key exists.

## 4. Corrected Fields

- `id_number_corrected`
- `passport_corrected`

## 5. Status Fields

- `identifier_status`
- `_validation_status` may include identifier findings from institution
  validation.

## 6. Original-Value Immutability Rules

Approved rule: source identifier fields must not be modified.

## 7. Corrected-Field Contract

- `id_number_corrected` contains the cleaned ID value that remains in the ID
  path. Valid 4-9 digit IDs are padded when appropriate. Numeric-only
  invalid-length IDs remain visible as their cleaned numeric value.
- `passport_corrected` contains cleaned passport content or empty string.
- Status text describes the ID/passport outcome.
- Numeric-only invalid IDs must remain in the ID path, must not be blanked, and
  must not be routed to passport only because of length.

## 8. Parsing / Cleanup / Normalization Rules

Approved rule:

- Passport cleanup keeps digits, ASCII letters, Hebrew letters, and dash
  variants. Other characters are removed.
- The ID value `9999` is treated as no ID.
- ID cleanup removes dash variants only. Other non-digit characters remain for
  classification.
- Any non-digit, non-dash character in the ID causes the ID value to be moved to
  passport when passport is empty.
- Numeric-only IDs are never moved to passport only because of length.
- IDs with 4 to 9 digits are left-padded to 9 digits for checksum validation.
- Numeric-only IDs shorter than 4 digits remain invalid IDs and are not moved to
  passport. They remain visible in `id_number_corrected` as cleaned numeric
  values.
- Numeric-only IDs longer than 9 digits remain invalid IDs and are not moved to
  passport. They remain visible in `id_number_corrected` as cleaned numeric
  values.
- All-zero and all-identical 9-digit IDs are invalid and are not moved.
- Israeli ID checksum uses alternating multipliers 1 and 2, subtracting 9 from
  doubled values greater than 9, and requiring a sum divisible by 10.

Current behavior:

- Hyphen-stripped valid IDs are exported from `id_number_corrected`.
- Invalid checksum IDs that are not moved may still remain in
  `id_number_corrected` as the normalized/padded numeric value.
- Numeric-only invalid-length IDs are treated as invalid IDs, not passports, and
  are not blanked from `id_number_corrected`.

## 9. Validation Rules

Institution validation:

- Requires a non-empty `id_number_corrected` only when the original ID is also
  missing; rejected IDs with an original value rely on `identifier_status`.
- Checks duplicates within a sheet using non-empty `id_number_corrected`.
- Checks duplicate IDs across workbook sheets and reports cross-sheet duplicate
  as warning.
- Empty corrected IDs are not considered duplicates.

## 10. Recovery Rules

On exception:

- `id_number_corrected` falls back to `id_number` if the source key exists.
- `passport_corrected` falls back to `passport` if the source key exists.
- both field names are added to `_standardization_failures` as applicable.
- `identifier_status` is set to empty string.

## 11. Ambiguity Rules

Current behavior:

- Passport-like ID values are moved only when they contain letters or special
  nonnumeric content other than allowed dash variants.
- If a passport already exists, an ID moved to passport does not overwrite that
  existing passport value.

Needs approval: no explicit conflict status exists for "ID looks like passport
but passport already had a value" beyond the current status text.

## 12. Invalid-Value Behavior

Invalid cases include:

- missing both ID and passport
- invalid checksum
- all-zero ID
- all-identical-digit ID
- too-short numeric-only ID
- too-long numeric-only ID
- non-digit/non-dash ID content

Invalid values write a Hebrew identifier status.

Rules:

- Values containing letters or special characters may be routed to
  `passport_corrected` if passport is empty.
- Numeric-only invalid IDs are not routed to passport and remain visible in
  `id_number_corrected`.
- Values moved to passport clear `id_number_corrected`.

## 13. Export Behavior

The active web export maps:

- `MisparZehut` from `id_number_corrected`
- `Darkon` from `passport_corrected`

The compatibility export can keep `passport_corrected` even when no source
passport column existed, which is covered by
`tests/test_export_engine_dataset.py`.

Export reads corrected values from the standardized dataset rows according to
the corrected-only export mapping. It must not depend on whether the UI/grid
shows a helper column.

## 14. UI/Grid Behavior

The UI places corrected ID/passport fields after their originals and anchors
`identifier_status` after the rightmost source field in the identifier group.

## 15. API Behavior

The standardization API builds `IdentifierEngine()` and enables identifier
standardization by default.

## 16. Examples

| Source ID | Source Passport | Corrected ID | Corrected Passport | Current Status |
|---|---|---|---|---|
| `000000018` | empty | `000000018` | empty | valid ID |
| `12345` | empty | padded 9-digit value | empty | valid/invalid by checksum |
| `ABC123` | empty | empty | `ABC123` | moved to passport |
| `123` | empty | `123` | empty | invalid short numeric ID |
| `12345678910` | empty | `12345678910` | empty | invalid long numeric ID |
| `AB@123` | empty | empty | `AB123` | moved to passport |
| empty | `ABC123` | empty | `ABC123` | passport entered |
| `9999` | empty | empty | empty | missing identifier |

## 17. Current Known Limitations

- Statuses are Hebrew display strings, not stable status codes.
- Invalid checksum values may remain in `id_number_corrected`.
- Existing passport values are not overwritten by moved ID values.

## 18. Open Questions Requiring Approval

- Should invalid checksum IDs be blanked from `id_number_corrected`?
- Should status codes be added beside Hebrew display statuses?
- Should conflict behavior be added when both passport and passport-like ID are
  present?

## 19. Tests That Should Cover The Behavior

- `tests/test_identifier_engine.py`
- `tests/test_export_engine_dataset.py`
- `tests/test_institution_report_validator.py`
- `tests/test_normalization_pipeline.py`

## 20. Final Principles

Identifiers are processed as a pair.

Do not validate ID and passport in separate passes unless the pairwise routing
rules are preserved.

Numeric-only invalid IDs remain ID values and must not automatically become
passport values only because of invalid length. They remain visible in
`id_number_corrected` with an invalid status.
