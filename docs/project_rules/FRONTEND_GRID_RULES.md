# Frontend Grid Rules

## Purpose

Document the browser grid as it exists now: a static JS UI that consumes the
API payloads shaped by the backend.

## Scope

- `webapp/templates/index.html`
- `webapp/static/js/grid.js`
- `webapp/static/js/edit.js`
- `webapp/static/js/upload.js`
- `webapp/static/js/export.js`
- `webapp/static/style.css`
- `webapp/services/workbook_service.py`

## Main Files

- `webapp/templates/index.html`
- `webapp/static/js/grid.js`
- `webapp/static/js/edit.js`
- `webapp/static/js/upload.js`
- `webapp/static/js/export.js`
- `webapp/static/style.css`
- `webapp/services/workbook_service.py`

## Responsibilities

- Render workbook sheet data in a table.
- Show corrected fields and status fields in the configured order.
- Support row selection, row deletion, and inline cell editing.
- Present the data without becoming the source of truth.

## Data Flow

1. Browser requests a sheet payload from the API.
2. `WorkbookService` returns shaped rows and field names.
3. Static JS renders the grid.
4. Edit/delete actions call the API and re-render from updated session state.

## Contracts

- `_row_uid` is the stable row identity in the UI.
- The grid is a helper view; it must not determine export correctness.
- Corrected columns should appear next to their source fields where the backend
  contract places them.

## What Must Never Change

- The grid must not invent or hide corrected values in a way that changes the
  underlying session row data.
- Delete/edit actions must continue to target `_row_uid`, not row indices.

## Current Behavior

- The UI is rendered by static JavaScript.
- Corrected and status columns are styled separately.
- Row deletion and inline edit actions operate on the session dataset.
- Backend grid payload shaping now lives in `webapp/services/grid_payload.py`.
- Shared grid metadata and source/corrected/status grouping live in
  `src/excel_standardization/normalized_row_contract.py`.

## Known Limitations

- There is no separate frontend build system in the repo.
- Grid payload shape is produced by backend shaping logic rather than a shared
  frontend model package.

## Tests That Should Cover It

- workbook service payload tests
- edit/delete regression tests
- browser/API integration tests for ordering and visibility

## Open Questions / Future Improvements

- Whether to split UI payload shaping from data loading in `WorkbookService`.
- Whether to add a dedicated frontend contract module.
