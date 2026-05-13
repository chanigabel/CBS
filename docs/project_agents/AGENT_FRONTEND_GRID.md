# Agent: Frontend Grid

## Mission

Review the browser grid, selection, editing, and visibility contract.

## Files To Inspect First

- `docs/project_rules/FRONTEND_GRID_RULES.md`
- `webapp/templates/index.html`
- `webapp/static/js/grid.js`
- `webapp/static/js/edit.js`
- `webapp/static/js/upload.js`
- `webapp/static/js/export.js`
- `webapp/static/style.css`
- `webapp/services/workbook_service.py`

## Rules To Follow

- Treat the grid as a helper view, not the source of truth.
- Keep `_row_uid` as the stable identity.

## What The Agent May Change

- JS/CSS payload handling, tests, and docs when requested.

## What The Agent Must Not Change

- Backend business rules.
- Export mapping.

## Required Tests

- UI payload ordering tests
- row-UID stability tests
- delete/edit regression tests

## Regression Checklist

- corrected columns appear in the intended order
- status columns remain visible
- delete/edit actions still use row UIDs

## Expected Output Format

- UI contract findings
- ordering/visibility findings
- tests run / missing

## Safety Constraints

- Do not let UI visibility control export behavior.
- Do not let filtering alter row identity.

