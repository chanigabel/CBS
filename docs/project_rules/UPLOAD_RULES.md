# Upload Rules

## Purpose

Describe how uploaded Excel files are accepted, stored, validated, and turned
into a session.

## Scope

- `webapp/services/upload_service.py`
- `webapp/api/upload.py`
- upload tests and fixture handling

## Main Files

- `webapp/services/upload_service.py`
- `webapp/api/upload.py`
- `tests/webapp/test_upload_service.py`
- `tests/webapp/test_api_upload.py`
- `tests/test_xls_legacy_support.py`

## Responsibilities

- Validate the uploaded filename extension.
- Save the original file and a working copy.
- Confirm the workbook can be opened.
- Create a session record with workbook paths.

## Data Flow

1. File bytes arrive with the browser request.
2. Extension is checked.
3. Source file is written to `uploads/`.
4. Working copy is written to `work/`.
5. Workbook opening is validated.
6. Sheet names are returned in the upload response.

## Contracts

- Original bytes are stored separately from the working copy.
- Supported extensions are `.xlsx`, `.xlsm`, and `.xls`.
- The upload flow must never mutate the source file.

## What Must Never Change

- The source file must remain untouched after upload.
- A failed upload must not leave behind a half-valid session.
- Unsupported extensions must fail clearly.

## Current Behavior

- Extension validation is suffix-based.
- `.xls` is validated through the legacy reader path.
- `.xlsx` / `.xlsm` are validated with openpyxl.

## Known Limitations

- Content sniffing is limited; the current gate is extension plus workbook open
  validation.
- Corrupt, locked, or password-protected files still depend on the reader
  library's exception behavior.

## Tests That Should Cover It

- accepted extension tests
- invalid extension tests
- corrupt workbook tests
- `.xls` and `.xlsx` acceptance tests

## Open Questions / Future Improvements

- Whether to add stronger MIME/content validation.
- Whether to surface more specific user-facing messages for corrupt/locked
  uploads.
