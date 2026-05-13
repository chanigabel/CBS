# Agent: API

## Mission

Review and maintain the FastAPI routers, request models, and HTTP responses.

## Files To Inspect First

- `docs/project_rules/API_RULES.md`
- `webapp/api/upload.py`
- `webapp/api/workbook.py`
- `webapp/api/edit.py`
- `webapp/api/institution.py`
- `webapp/api/export.py`
- `webapp/models/requests.py`
- `webapp/models/responses.py`

## Rules To Follow

- Keep HTTP contracts aligned with the real services.
- Do not move business logic into routers.

## What The Agent May Change

- Endpoint wiring, response models, tests, and docs when requested.

## What The Agent Must Not Change

- Standardization business rules.
- Session state semantics.
- Export mapping policy.

## Required Tests

- API upload tests
- workbook summary/sheet tests
- edit/delete tests
- normalize/export endpoint tests

## Regression Checklist

- Upload returns a usable session and sheet list.
- Workbook endpoints return the shaped UI payload.
- Edit and delete endpoints remain row-UID based.
- Export endpoint returns a downloadable workbook.

## Expected Output Format

- Endpoint-level findings
- response and error behavior
- tests that cover the issue
- tests that are missing

## Safety Constraints

- Preserve source data immutability.
- Keep status and corrected-field contracts intact.

