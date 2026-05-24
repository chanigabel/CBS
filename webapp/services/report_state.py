"""Helpers for keeping report-related session state in sync."""

from __future__ import annotations

from copy import deepcopy
from typing import Any

from webapp.services.row_identity import row_lookup

SOURCE_ROW_COUNT_KEY = "source_row_count"


def ensure_source_row_count(sheet) -> None:
    """Persist the original row count on a sheet once, if missing."""
    if sheet.get_metadata(SOURCE_ROW_COUNT_KEY) is None:
        sheet.set_metadata(SOURCE_ROW_COUNT_KEY, len(sheet.rows))


def source_row_count(sheet) -> int:
    """Return the original source row count for a sheet."""
    value = sheet.get_metadata(SOURCE_ROW_COUNT_KEY)
    try:
        return int(value)
    except (TypeError, ValueError):
        return len(sheet.rows)


def snapshot_workbook_dataset(workbook_dataset):
    """Return a deep copy of the workbook dataset for report baseline tracking."""
    return deepcopy(workbook_dataset)


def _normalized_value(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def sync_edit_tracking(record, sheet_name: str, row_uid: str, field_name: str, new_value: Any) -> None:
    """Keep the session's edit tracking aligned with the standardized baseline.

    A cell only remains tracked as a manual edit while its current value differs
    from the standardized baseline snapshot.
    """
    baseline = getattr(record, "report_baseline_workbook_dataset", None)
    key = (sheet_name, row_uid, field_name)

    if baseline is None:
        record.edits[key] = new_value
        return

    sheet = baseline.get_sheet_by_name(sheet_name)
    if sheet is None:
        record.edits[key] = new_value
        return

    baseline_row = row_lookup(sheet).get(row_uid)
    if baseline_row is None:
        record.edits[key] = new_value
        return

    _idx, row = baseline_row
    baseline_value = row.get(field_name)
    if _normalized_value(new_value) == _normalized_value(baseline_value):
        record.edits.pop(key, None)
    else:
        record.edits[key] = new_value


def remove_edits_for_row_uids(record, sheet_name: str, row_uids: list[str]) -> None:
    """Remove edit-tracking entries for deleted rows."""
    if not row_uids:
        return
    uid_set = {str(uid) for uid in row_uids}
    keys_to_remove = [
        key for key in record.edits
        if isinstance(key, tuple) and len(key) == 3
        and str(key[0]) == str(sheet_name)
        and str(key[1]) in uid_set
    ]
    for key in keys_to_remove:
        record.edits.pop(key, None)
