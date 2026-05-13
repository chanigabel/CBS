# Agent: Project Review

## Mission

Perform a high-signal review of the whole project: architecture, API, upload,
workbook/session flow, extraction, pipeline, validation, export, UI/grid,
manual edits, packaging, testing, security, and legacy paths.

## Files To Inspect First

- `docs/project_rules/README.md`
- `docs/project_rules/PROJECT_ARCHITECTURE_RULES.md`
- `docs/project_rules/WORKBOOK_LOADER_RULES.md`
- `docs/project_rules/STANDARDIZATION_PIPELINE_RULES.md`
- `docs/project_rules/EXPORT_SYSTEM_RULES.md`
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
- Treat `webapp/services/workbook_loader.py` as the canonical workbook
  dispatch path.
- Treat `src/excel_standardization/normalized_row_contract.py` as the shared
  corrected-field / export-field contract helper.
- Treat `webapp/services/grid_payload.py` as the backend grid payload builder.
- Treat export assembly as row-view based and non-mutating.

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
