# Testing Rules

## Purpose

Describe the regression strategy used to keep the project stable.

## Scope

- `tests/`
- pytest configuration and project-specific fixtures

## Main Files

- `tests/`
- `pyproject.toml`
- `README.md`

## Responsibilities

- Cover business rules, UI/data contracts, and workbook IO.
- Lock down regression bugs with focused tests.
- Separate fast unit tests from broader integration tests.

## Data Flow

1. Unit tests verify engine and helper behavior.
2. Service tests verify session and API flows.
3. Integration tests verify upload -> normalize -> export paths.
4. Regression tests pin the bugs that have already been fixed.

## Contracts

- Tests should assert current behavior, not imagined behavior.
- Any change to corrected-field flow, UI visibility, or export output should
  receive a targeted regression test.

## What Must Never Change

- Approved business rules should be locked by tests.
- Regressions in workbook openability or row identity must be caught early.

## Current Behavior

- The repository has separate test coverage for engines, pipeline, web
  services, API routes, and export output.

## Known Limitations

- Some behaviors are only covered at service level, not fully end-to-end.
- A few areas still rely on coordinated unit tests rather than a shared
  contract test.

## Tests That Should Cover It

- engine tests
- pipeline tests
- workbook/session tests
- export tests
- upload/load tests
- manual edit/delete tests

## Open Questions / Future Improvements

- Whether to introduce a small number of end-to-end smoke tests that span all
  major services.
- Whether to standardize a naming convention for regression tests.

