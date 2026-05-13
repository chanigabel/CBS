# Legacy / Disabled Paths Rules

## Purpose

Document the archived or disabled paths that remain in the repository for
reference.

## Scope

- `src/excel_standardization/orchestrator.py`
- `archive_legacy/`
- any disabled direct-workbook path helpers

## Main Files

- `src/excel_standardization/orchestrator.py`
- `archive_legacy/`

## Responsibilities

- Preserve historical behavior for reference.
- Keep the active runtime on the dataset/session pipeline.

## Data Flow

1. Legacy entry points remain importable for compatibility.
2. Public methods raise a disabled-path error.
3. The active web pipeline handles all real processing.

## Contracts

- Disabled paths should stay disabled unless there is a deliberate reactivation
  plan.
- Documentation should make it clear that the active runtime is the
  web/session pipeline.

## What Must Never Change

- Legacy direct-Excel processing must not quietly become the default again.
- Disabled public methods must keep failing clearly.

## Current Behavior

- `StandardizationOrchestrator` raises a runtime error for legacy direct
  workbook methods.
- The file remains as a facade for historical compatibility.

## Known Limitations

- Legacy code and comments still exist because the project kept them for audit
  and reference.

## Tests That Should Cover It

- tests that assert disabled methods fail
- tests that verify the active web/session pipeline still works

## Open Questions / Future Improvements

- Whether to eventually remove the legacy facade entirely.
- Whether to relocate all archived notes into one clearly labeled archive tree.

