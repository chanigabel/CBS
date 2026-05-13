# Gender Rules

## 1. Purpose

Normalize gender values to the export codes `1` and `2` while preserving the
original input and surfacing invalid values through a status field.

## 2. Scope

Applies to the `gender` source field.

Implemented by:

- `src/excel_standardization/engines/gender_engine.py`
- `src/excel_standardization/processing/gender_standardization.py`
- `src/excel_standardization/validation/institution_report_validator.py`

## 3. Source Fields

- `gender`

## 4. Corrected Fields

- `gender_corrected`

## 5. Status Fields

- `gender_status`
- `_validation_status` may also include gender-related validation findings in
  institution-report validation.

## 6. Original-Value Immutability Rules

Approved rule: `gender` must not be modified. All normalized values go into
`gender_corrected`.

## 7. Corrected-Field Contract

- Recognized male values produce integer `1`.
- Recognized female values produce integer `2`.
- Unrecognized non-empty values produce an empty string in `gender_corrected`.
- Missing values and `None` remain empty/blank in `gender_corrected`.
- Empty string values remain empty in `gender_corrected`.
- Whitespace-only values normalize to an empty string in `gender_corrected`.
- The original `gender` value is always preserved unchanged.

## 8. Parsing / Cleanup / Normalization Rules

Approved rule:

- Convert non-empty values to a stripped lowercase string.
- Check female patterns before male patterns.
- Use substring matching, not exact matching.
- Female patterns include numeric `2`, English female terms, and configured
  Hebrew female terms.
- Male patterns include numeric `1`, English male terms, and configured Hebrew
  male terms.

Current behavior: because matching is substring-based, a longer value containing
one of the configured tokens may normalize even if it is not an exact gender
field value.

Approved decision: keep substring matching for now.

## 9. Validation Rules

Institution validation accepts corrected values `1` and `2`.

- Empty `gender_corrected` from an unrecognized value is invalid.
- Missing gender fields are allowed by the validator.
- Original present with no corrected value is treated as unvalidated and skipped
  by the validator.

## 10. Recovery Rules

On exception:

- `gender_corrected` falls back to the original value.
- `gender` is listed in `_standardization_failures`.

## 11. Ambiguity Rules

Approved rule: female patterns are checked first so values like `female` do not
match the male `m` pattern first.

Current behavior: there is no separate ambiguous-gender status.

## 12. Invalid-Value Behavior

Invalid non-empty values:

- produce `gender_corrected = ""`
- write `gender_status` with a Hebrew invalid-code message
- are also flagged by institution validation when applicable

Whitespace-only values:

- produce `gender_corrected = ""`
- should not preserve spaces in the corrected field
- should preserve the original `gender` source value unchanged

## 13. Export Behavior

All export paths should use `gender_corrected` for the standardized `Min`
column.

Approved export rule:

- `Min` is exported from `gender_corrected`.
- Export must not fall back to the original `gender` value when
  `gender_corrected` exists but is empty.
- If `gender_corrected` is missing because standardization was not run, the
  standardized export value should remain blank rather than exporting an
  unstandardized raw gender value.

## 14. UI/Grid Behavior

The UI places `gender_corrected` immediately after `gender` when present, and
places `gender_status` after the corrected gender column.

## 15. API Behavior

The standardization API builds `GenderEngine()` and enables gender
standardization by default.

## 16. Examples

| Source | Corrected | Status |
|---|---:|---|
| `1` | `1` | empty |
| `male` | `1` | empty |
| `2` | `2` | empty |
| `female` | `2` | empty |
| `8` | empty string | invalid gender code |
| `xyz` | empty string | invalid gender code |
| whitespace only | empty string | empty |

## 17. Current Known Limitations

- Matching is substring-based by approved current behavior.
- There is no stable machine-readable gender status code.
- Stable status-code support is a future architectural improvement, not part of
  the current implementation requirement.

## 18. Open Questions Requiring Approval

- Should gender matching become exact-token matching in the future?
- Should invalid gender use a stable status code in addition to Hebrew text in
  a future cross-system status registry?

## 19. Tests That Should Cover The Behavior

- `tests/test_gender_engine.py`
- `tests/test_institution_report_validator.py`
- `tests/test_normalization_pipeline.py`
- web grid tests for status placement
- export tests proving `Min` is written from `gender_corrected` only
- tests proving whitespace-only gender values produce `gender_corrected = ""`

## 20. Final Principles

Gender normalization must never copy invalid raw values into `gender_corrected`.
Only approved codes `1` and `2` are valid corrected values.

Original `gender` values remain immutable. Corrected/exported values must be
standardized or blank.