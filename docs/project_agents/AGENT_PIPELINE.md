# Agent: Standardization Pipeline

## Mission

Review and maintain the end-to-end normalization pipeline and its web service
integration.

## Files To Inspect First

- `docs/project_rules/STANDARDIZATION_PIPELINE_RULES.md`
- `docs/project_rules/WORKBOOK_LOADER_RULES.md`
- `src/excel_standardization/processing/standardization_pipeline.py`
- `src/excel_standardization/processing/name_standardization.py`
- `src/excel_standardization/processing/gender_standardization.py`
- `src/excel_standardization/processing/date_standardization.py`
- `src/excel_standardization/processing/identifier_standardization.py`
- `webapp/services/standardization_service.py`

## Rules To Follow

- Keep original values immutable.
- Keep corrected fields and statuses visible to downstream consumers.
- Treat `src/excel_standardization/normalized_row_contract.py` as the shared
  source for corrected-field selection and export-field selection.

## What The Agent May Change

- Orchestration, metadata/stats, tests, and docs when requested.

## What The Agent Must Not Change

- Approved engine behavior.
- Export mapping.
- UI row identity semantics.

## Required Tests

- pipeline unit tests
- normalization service tests
- engine regression tests
- workbook-level validation after normalization

## Regression Checklist

- single-sheet normalization
- full-workbook normalization
- partial failure isolation
- manual edit replay after normalization

## Expected Output Format

- stage-order findings
- metadata/stats findings
- UI/export impact
- tests run / missing

## Safety Constraints

- Do not rewrite the whole pipeline.
- Keep the active web session flow working.
