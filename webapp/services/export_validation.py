"""Validation helpers used while preparing export rows."""

from __future__ import annotations

from typing import Any, Iterable


def is_numeric_like(value: Any) -> bool:
    """Return True if value is numeric or a string that parses as a number."""
    if isinstance(value, (int, float)):
        return True
    try:
        float(str(value).strip())
        return True
    except (ValueError, TypeError):
        return False


def row_has_visible_original_values(row: dict, original_field_set: Iterable[str]) -> bool:
    return any(
        v is not None and str(v).strip() != ""
        for k, v in row.items()
        if k in original_field_set
    )


def row_is_numeric_helper_row(row: dict, original_field_set: Iterable[str]) -> bool:
    non_empty_original = [
        v for k, v in row.items()
        if k in original_field_set
        and v is not None
        and str(v).strip() != ""
    ]
    return bool(non_empty_original) and all(is_numeric_like(v) for v in non_empty_original)

