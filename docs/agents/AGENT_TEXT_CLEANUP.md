# Agent Text Cleanup

## 1. Mission

Review and maintain low-level text cleanup behavior used by name
standardization.

## 2. Files To Inspect First

- `docs/standardization_rules/TEXT_CLEANUP_RULES.md`
- `src/excel_standardization/engines/text_processor.py`
- `src/excel_standardization/engines/name_engine.py`
- `tests/test_name_engine.py`

## 3. Rules Documents To Follow

- `docs/standardization_rules/TEXT_CLEANUP_RULES.md`
- `docs/standardization_rules/NAME_RULES.md`

## 4. What The Agent May Change

- Text cleanup helpers and focused tests when requested.
- Token lists only when the new token behavior is approved.

## 5. What The Agent Must Not Change

- Cleanup ordering without explicit review.
- `clean_text` alias behavior without approval.
- Original values in callers.

## 6. Required Safety Constraints

- Cleanup must be deterministic.
- Removed content must not be inferred or replaced with guessed content.
- If cleanup removes all text, return empty string.

## 7. Required Tests Before/After Changes

- `pytest tests/test_name_engine.py`

## 8. Expected Output Format

List changed cleanup step, before/after examples, and tests run.

## 9. Review Checklist

- Language detection happens before character filtering.
- Parenthesized acronym removal happens before filtering.
- Unwanted-token removal happens after filtering.
- Hyphen/backslash behavior remains covered.

## 10. Regression Checklist

- Hebrew-only names.
- English-only names.
- Mixed-script names.
- Diacritics.
- Title tokens.
- Parentheses.
