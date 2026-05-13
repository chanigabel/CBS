# Agent: Workbook / Session

## Mission

Review and maintain the session-backed workbook state, summaries, sheet access,
and manual edit replay behavior.

## Files To Inspect First

- `docs/project_rules/WORKBOOK_SESSION_RULES.md`
- `webapp/services/session_service.py`
- `webapp/models/session.py`
- `webapp/services/workbook_service.py`
- `webapp/services/standardization_service.py`
- `webapp/services/edit_service.py`

## Rules To Follow

- Keep `_row_uid` stable.
- Keep workbook state in the session.
- Do not let row filtering or deletion change identity semantics.

## What The Agent May Change

- Workbook/session helpers, tests, and docs when requested.

## What The Agent Must Not Change

- Original field values.
- Corrected-field creation logic.
- Export mapping policy.

## Required Tests

- workbook summary tests
- sheet data tests
- edit/delete tests
- replay-after-standardization tests

## Regression Checklist

- sheet loading is lazy and stable
- `_row_uid` survives repeated calls
- manual edits replay after normalization
- row deletion stays row-UID based

## Expected Output Format

- Session flow findings
- API payload findings
- risks
- tests run and missing

## Safety Constraints

- Treat the session dataset as the active working state.
- Do not silently reindex rows.

