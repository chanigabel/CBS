# Agent: Project Review

## Mission

Perform a high-signal review of the whole project: architecture, API, upload,
workbook/session flow, extraction, pipeline, validation, export, UI/grid,
manual edits, packaging, testing, security, and legacy paths.

## Files To Inspect First

- `docs/project_rules/README.md`
- `docs/project_rules/PROJECT_ARCHITECTURE_RULES.md`
- `webapp/app.py`
- `webapp/api/*.py`
- `webapp/services/*.py`
- `src/excel_standardization/`
- `tests/`

## Rules To Follow

- Use the active codebase as the first source of truth.
- Do not invent behavior.
- Separate approved behavior, current behavior, known limitations, and needs
  approval.

## What The Agent May Change

- Documentation and tests when explicitly asked.

## What The Agent Must Not Change

- Runtime code unless the user explicitly asks for implementation.
- Approved business rules.

## Required Tests

- Relevant tests for each touched subsystem.

## Regression Checklist

- Upload still accepts supported Excel formats.
- Workbook/session flow still preserves row identity.
- Export still opens in Excel.
- Manual edits still replay after normalization.

## Expected Output Format

- Summary
- Files inspected
- Findings by subsystem
- Tests covered / missing
- Risks and follow-up items

## Safety Constraints

- Treat source workbooks as immutable.
- Keep corrected-only export behavior intact.
- Keep UI payloads separate from source data contracts.

