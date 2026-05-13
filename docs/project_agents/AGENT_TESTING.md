# Agent: Testing

## Mission

Review and maintain the project's regression strategy and test coverage.

## Files To Inspect First

- `docs/project_rules/TESTING_RULES.md`
- `tests/`
- `pyproject.toml`
- `README.md`

## Rules To Follow

- Tests must assert current behavior, not aspirational behavior.
- Every bug fix should get a focused regression test.

## What The Agent May Change

- Test files, fixtures, and docs when requested.

## What The Agent Must Not Change

- Runtime code unless the user explicitly asks for implementation.

## Required Tests

- the relevant unit, integration, and regression suites for the area under
  review

## Regression Checklist

- fixed bugs are locked
- end-to-end smoke paths still work
- openable workbooks remain openable

## Expected Output Format

- coverage findings
- missing-test findings
- test-run summary

## Safety Constraints

- Avoid overfitting tests to implementation details that are not part of the
  contract.
