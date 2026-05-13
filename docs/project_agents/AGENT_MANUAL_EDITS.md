# Agent: Manual Edits

## Mission

Review and maintain cell edits, row deletions, and edit replay behavior.

## Files To Inspect First

- `docs/project_rules/MANUAL_EDITS_RULES.md`
- `webapp/services/edit_service.py`
- `webapp/api/edit.py`
- `webapp/static/js/edit.js`
- `webapp/models/requests.py`
- `tests/webapp/test_edit_service.py`
- `tests/webapp/test_hotfix_row_shift.py`

## Rules To Follow

- Use `_row_uid` for all row-level operations.
- Preserve the user's manual edits through normalization replay.

## What The Agent May Change

- Edit/delete helpers, tests, and docs when requested.

## What The Agent Must Not Change

- Source workbook data.
- Standardization business rules.

## Required Tests

- edit service tests
- row deletion tests
- replay-after-standardization tests

## Regression Checklist

- edited cell values persist in session state
- deleted rows disappear from the current dataset
- no row-shift regressions

## Expected Output Format

- edit flow findings
- deletion flow findings
- replay findings
- tests run / missing

## Safety Constraints

- Never fall back to row indexes.
- Keep the session data model stable.
