# API Rules

## Purpose

Document the current HTTP contract so agents can change endpoints without
breaking request/response behavior.

## Scope

- `webapp/api/*.py`
- request/response models
- HTTP status handling
- file upload and workbook/session endpoints
- edit, delete, normalize, and export routes

## Main Files

- `webapp/api/upload.py`
- `webapp/api/workbook.py`
- `webapp/api/edit.py`
- `webapp/api/institution.py`
- `webapp/api/export.py`
- `webapp/models/requests.py`
- `webapp/models/responses.py`

## Responsibilities

- Parse HTTP input.
- Call the appropriate service.
- Return a stable response model or HTTP error.
- Never embed business rules that belong in engines or validators.

## Data Flow

1. Request enters FastAPI.
2. Router validates the request model.
3. Service mutates or reads session-backed state.
4. Response model or file response returns to the browser/client.

## Contracts

- API endpoints must not mutate source files.
- IDs, row UIDs, and session IDs are the stable identifiers.
- Errors should be explicit enough to help the user fix the input.

## What Must Never Change

- The API should not rely on UI visibility to determine data correctness.
- Export should not depend on whether a field is visible in the grid.
- Manual edits and deletions must continue to use `_row_uid`.

## Current Behavior

- Upload accepts Excel files and creates a session.
- Workbook endpoints expose summaries and sheet rows.
- Edit endpoints mutate the in-memory session dataset.
- Export endpoints generate a downloadable workbook from the session.

## Known Limitations

- Error handling is still partly generalized around invalid workbook input.
- Some endpoint docstrings/comments may lag behind current file-type support.

## Tests That Should Cover It

- router/API tests for upload, workbook, edit, validation, and export
- upload/export happy-path and error-path tests

## Open Questions / Future Improvements

- Whether to standardize error codes more aggressively across workbook load
  failures.
- Whether to add a shared API error contract for workbook-related exceptions.

