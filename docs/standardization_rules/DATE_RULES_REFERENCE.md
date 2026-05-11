# Date Rules Reference

## Purpose

This file keeps the organized standardization documentation tree complete without
rewriting the existing date rules document.

## Scope

Date behavior is governed by the root-level [`DATE_RULES.md`](../../DATE_RULES.md).
That document already contains the detailed contract for:

- source fields and corrected fields
- component-based date export
- original-value immutability
- parsing, recovery, ambiguity, and invalid-date behavior
- visible status behavior
- compact numeric dates
- Excel serial dates
- split-date and single-date behavior
- list-level birth-year majority correction

## Active Implementation Files

- `src/excel_standardization/engines/date_engine.py`
- `src/excel_standardization/processing/date_standardization.py`
- `src/excel_standardization/processing/standardization_pipeline.py`
- `src/excel_standardization/export/export_engine.py`
- `webapp/services/export_schema.py`
- `webapp/services/workbook_service.py`
- `webapp/services/export_writer.py`

## Tests To Inspect

- `tests/test_date_engine.py`
- `tests/test_date_engine_corrected_flow.py`
- `tests/test_date_conservative_parsing.py`
- `tests/test_parse_date_orchestration.py`
- `tests/test_plain_date_columns.py`
- `tests/test_per_field_date_detection.py`
- `tests/webapp/test_compact_date_extraction.py`

## Current Behavior Summary

- Date standardization writes structured corrected components:
  `birth_year_corrected`, `birth_month_corrected`, `birth_day_corrected`,
  `entry_year_corrected`, `entry_month_corrected`, and `entry_day_corrected`.
- The active export schema maps date export columns to corrected components only.
- Date statuses are written to `birth_date_status` and `entry_date_status`.
- Internal date helper keys start with `_` and must not be shown in the UI grid.
- The UI groups original date fields, corrected date components, and the related
  status column together.

## Needs Approval

- Do not duplicate or fork the root `DATE_RULES.md` unless the project chooses
  to move the source of truth into `docs/standardization_rules/`.
- If moved later, all references in agent docs and tests should be updated in
  the same change.

## Final Principles

Date rules remain the most detailed rule set in the project. Any future date
change must be reviewed against `DATE_RULES.md`, the date tests, UI grid
behavior, and export behavior together.
