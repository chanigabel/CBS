# Manual Edits Rules

## Purpose

Document how users edit cells and delete rows in the session dataset.

## Scope

- `webapp/services/edit_service.py`
- `webapp/api/edit.py`
- `webapp/static/js/edit.js`
- `tests/webapp/test_edit_service.py`
- `tests/webapp/test_hotfix_row_shift.py`

## Main Files

- `webapp/services/edit_service.py`
- `webapp/api/edit.py`
- `webapp/static/js/edit.js`
- `tests/webapp/test_edit_service.py`
- `tests/webapp/test_hotfix_row_shift.py`

## Responsibilities

- Edit one cell by row UID.
- Delete one or more rows by row UID.
- Preserve a stable edit/deletion target across re-rendering.
- Replay edits after standardization when the session is normalized again.

## Data Flow

1. User edits or deletes from the grid.
2. Browser sends row UID-based API calls.
3. `EditService` mutates the in-memory dataset.
4. The session records edits for replay.

## Contracts

- `_row_uid` is the stable target for edits and deletions.
- Edits are recorded under `(sheet_name, row_uid, field_name)`.
- Deletion removes the row from the current in-memory dataset.

## What Must Never Change

- Edits must not fall back to row index.
- Deleted rows must not be re-targeted accidentally after filtering or sorting.
- Replay must preserve the user's chosen cell value.

## Current Behavior

- Cell edits coerce numeric values back to the original type where possible.
- Row deletions update the session dataset in memory.
- Standardization replays stored edits after normalization.

## Known Limitations

- Deleted rows are not recoverable unless the user reruns from the source file.
- The session stores edits in memory rather than a separate persistent audit log.

## Tests That Should Cover It

- edit service unit tests
- row-shift regression tests
- replay-after-standardization tests

## Open Questions / Future Improvements

- Whether to record deletions as first-class session history.
- Whether to expose edit metadata in a more structured audit trail.
