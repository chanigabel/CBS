# Workbook / Session Rules

## Purpose

Document how workbook state is stored in memory and tied to a session.

## Scope

- `webapp/services/session_service.py`
- `webapp/models/session.py`
- `webapp/services/workbook_service.py`
- `webapp/services/standardization_service.py`
- `webapp/services/edit_service.py`

## Main Files

- `webapp/services/session_service.py`
- `webapp/models/session.py`
- `webapp/services/workbook_service.py`
- `webapp/services/standardization_service.py`
- `webapp/services/edit_service.py`

## Responsibilities

- Hold the workbook dataset in the session.
- Track edits, status, and processing reports.
- Provide workbook summaries and sheet data.
- Preserve row identity across UI actions.

## Data Flow

1. Upload creates a session record.
2. Workbook data is loaded lazily or on demand.
3. The session keeps the mutable in-memory dataset.
4. UI actions read from and write to the session.
5. Standardization and export reuse the same session state.

## Contracts

- `WorkbookDataset` belongs to the session, not the request.
- `_row_uid` must remain stable across reads and edits.
- Manual edits are stored separately from source extraction.

## What Must Never Change

- Session data must not silently switch to a different workbook.
- Deleting a row must be based on `_row_uid`, not row position.
- Manual edits must survive standardization replay.

## Current Behavior

- Workbook data can start as `None` and be loaded lazily.
- The session carries the current working copy paths and workbook dataset.
- Edits are recorded in `record.edits`.

## Known Limitations

- Session state is in-memory; durability depends on the process lifetime.
- There is no separate persistence layer for edits or workbook datasets.

## Tests That Should Cover It

- workbook/session summary tests
- edit/delete tests
- standardization replay tests
- lazy-load tests

## Open Questions / Future Improvements

- Whether session state should be persisted externally.
- Whether edit history should be surfaced as a first-class audit log.

