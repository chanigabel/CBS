# Agent: Legacy Paths

## Mission

Review archived or disabled code paths and keep them clearly separated from the
active runtime.

## Files To Inspect First

- `docs/project_rules/LEGACY_DISABLED_PATHS_RULES.md`
- `src/excel_standardization/orchestrator.py`
- `archive_legacy/`
- related tests that assert disabled behavior

## Rules To Follow

- The active runtime is the web/session pipeline.
- Disabled methods must keep failing clearly.

## What The Agent May Change

- Documentation and disabled-path tests when requested.

## What The Agent Must Not Change

- Re-enable legacy direct workbook processing without explicit approval.

## Required Tests

- tests for disabled methods raising errors
- tests proving the active web/session flow still works

## Regression Checklist

- legacy entry points remain disabled
- archive notes stay labeled as historical

## Expected Output Format

- disabled-path findings
- compatibility findings
- test coverage gaps

## Safety Constraints

- Do not confuse archived code with the active contract.
- Do not accidentally route runtime users through the legacy path.

