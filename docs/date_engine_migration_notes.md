# DateEngine Migration Notes

## What Changed

The active DateEngine flow now uses a structured, deterministic parsing architecture while preserving existing public APIs and field names.

The main behavior changes are:

- partial split dates remain split dates
- split and single date flows use the same corrected-field policy
- all two-digit year expansions carry metadata
- DateEngine receives an explicit processing date/reference date from the pipeline/service boundary
- majority correction can include single-date two-digit inputs
- Dataset export no longer falls back from corrected date fields to original date fields

## Backward Compatibility

The following are preserved:

- original raw fields
- corrected field names
- status field names
- Web export schema names
- DateEngine public wrapper methods
- current Hebrew status texts where existing behavior depended on them

Legacy callers can still call:

```python
engine.parse_date(...)
engine.parse_from_split_columns(...)
engine.parse_date_value(...)
engine.parse_numeric_date_string(...)
engine.parse_separated_date_string(...)
```

New active runtime code should prefer `DateInput` through `parse_input()`.

## Behavioral Differences To Expect

### Partial Split Dates

Before:

Partial split rows could fall through as empty or main-value dates.

After:

Any split component keeps the row in split mode. Valid components are preserved, missing components are blanked in corrected fields, and date status explains the missing component.

### Invalid Single Dates

Before:

Single-date flow could write raw parsed components directly.

After:

Single and split flows both use the shared corrected-field builder.

### Two-Digit Years

Before:

Some parser paths expanded years without setting `year_was_auto_completed`.

After:

All two-digit year parser paths set auto-completion metadata.

### Processing Date / Reference Date

Before:

Low-level parsing and business validation could call `date.today()` directly in multiple places.

After:

The service/pipeline captures one processing date per standardization run. It defaults to `date.today()`, remains business-significant, and DateEngine uses that same captured date consistently.

The normalized dataset metadata includes:

```python
processing_date
processing_year
```

### Export

Before:

Dataset export could fall back from corrected date fields to original date fields.

After:

Dataset export and Web export use corrected date fields only.

## Migration Risks

1. Files containing plain integer years may behave differently from true Excel serial dates.
2. More rows may be eligible for majority correction because metadata is now complete.
3. Existing downstream consumers that expected raw-date fallback in Dataset export will now see blank corrected values instead.
4. Tests or UI code that compare exact Hebrew status text should gradually move to status-code assertions where possible.
5. Partial split rows now surface explicit missing-component statuses, which may expose data quality issues that were previously hidden.
6. Processing-date behavior is intentionally year-sensitive; tests should pin the processing date when asserting exact year outcomes.

## Compatibility Strategy

- Keep all public DateEngine methods.
- Keep existing row field names.
- Keep raw values untouched.
- Add structured fields without removing legacy fields.
- Keep export headers unchanged.
- Keep Web/Dataset flow as the source of truth.

## Recommended Follow-Up Migration

1. Add extraction metadata for Excel cell date formats.
2. Use that metadata before treating integer values as Excel serial dates.
3. Add UI support for status codes if needed.
4. Decide whether age over 100 is warning-only or blocking validation.
5. Review whether the 1906 minimum year should apply to entry dates or birth dates only.
6. Decide whether majority correction should remain one-way.
