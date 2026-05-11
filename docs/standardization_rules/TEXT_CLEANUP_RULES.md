# Text Cleanup Rules

## 1. Purpose

Define the generic text cleanup behavior used by name standardization.

## 2. Scope

Applies to `TextProcessor.clean_name` and the legacy alias
`TextProcessor.clean_text`.

Implemented by:

- `src/excel_standardization/engines/text_processor.py`
- `src/excel_standardization/engines/name_engine.py`

## 3. Source Fields

Current behavior: text cleanup is used for name fields. It is not a general
standardization pass over every workbook cell.

## 4. Corrected Fields

Text cleanup writes no fields directly. Callers write corrected fields such as:

- `first_name_corrected`
- `last_name_corrected`
- `father_name_corrected`

## 5. Status Fields

Current behavior: text cleanup writes no status fields.

## 6. Original-Value Immutability Rules

Approved rule: callers must preserve original source fields and write cleaned
output to corrected fields.

## 7. Corrected-Field Contract

`clean_name` returns a string. `None`, unconvertible values, and values that
clean down to no accepted characters return an empty string.

## 8. Cleanup / Normalization Rules

Approved fixed order:

1. Convert safely to string.
2. Remove zero-width/invisible characters.
3. Remove configured diacritics.
4. Translate Arabic-Indic digits.
5. Detect language dominance.
6. Remove parenthesized acronym groups when the group contains a quote/acronym
   character.
7. Filter characters by dominant language.
8. Convert hyphen-like characters, parentheses, and backslash to spaces.
9. Drop digits, symbols, and wrong-language letters.
10. Collapse spaces.
11. Remove unwanted Hebrew or English title tokens.

Language dominance:

- Hebrew letters and English letters are counted.
- Hebrew wins ties.
- No Hebrew and no English letters results in `MIXED`.

## 9. Validation Rules

Text cleanup itself does not validate business requirements. Validators consume
the corrected fields produced by callers.

## 10. Recovery Rules

Current behavior: the helper returns empty string for `None` or empty input.
Exception recovery is handled by caller modules.

## 11. Ambiguity Rules

Current behavior: mixed Hebrew/English strings are resolved through dominance.
Tie goes to Hebrew. This can remove English characters from equal mixed input.

## 12. Invalid-Value Behavior

Values containing only digits, unsupported symbols, or removed tokens can produce
an empty corrected result. No status is written by text cleanup.

## 13. Export Behavior

Export does not call `TextProcessor` directly. Export receives values already
written by the pipeline.

## 14. UI/Grid Behavior

The UI shows text cleanup only through corrected fields written by the name
pipeline.

## 15. API Behavior

The web standardization service builds one `TextProcessor` and passes it to
`NameEngine`.

## 16. Examples

| Input | Output |
|---|---|
| `Smith-Jones` | `Smith Jones` |
| Hebrew name with digits | Hebrew letters only |
| Hebrew name with backslash between parts | Hebrew parts separated by a space |
| Parenthesized acronym containing quote | group removed |
| Normal Hebrew word in parentheses | word kept, parentheses removed |
| English name with accented letters | accent removed |

## 17. Current Known Limitations

- Diacritic mapping is a configured map, not full Unicode transliteration.
- Language dominance is count-based and does not detect names semantically.
- `fix_hebrew_final_letters` exists and is tested directly but is not part of
  the main `clean_name` pipeline.

## 18. Open Questions Requiring Approval

- Should mixed Hebrew/English equal-count values keep both scripts instead of
  Hebrew winning ties?
- Should text cleanup write a status when it removes all content?
- Should `clean_text` remain an alias for `clean_name`, or should generic text
  cleanup diverge from name cleanup?

## 19. Tests That Should Cover The Behavior

- `tests/test_name_engine.py`

## 20. Final Principles

Text cleanup is deterministic character processing. It must not infer missing
text or modify original fields.
