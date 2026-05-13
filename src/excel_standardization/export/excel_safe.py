"""Safety helpers for writing openpyxl workbooks."""

from __future__ import annotations

import json
import logging
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from typing import Any, Iterable, Set

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

logger = logging.getLogger(__name__)

_INVALID_SHEET_CHARS = set("[]:*?/\\")
_MAX_SHEET_TITLE_LEN = 31


def safe_sheet_title(title: Any, used_titles: Iterable[str] = ()) -> str:
    """Return an Excel-safe, unique worksheet title."""
    raw = "" if title is None else str(title)
    cleaned = "".join("_" if ch in _INVALID_SHEET_CHARS else ch for ch in raw).strip()
    cleaned = ILLEGAL_CHARACTERS_RE.sub("", cleaned) or "Sheet"

    used: Set[str] = set(used_titles)
    base = cleaned[:_MAX_SHEET_TITLE_LEN] or "Sheet"
    candidate = base
    counter = 1
    while candidate in used:
        suffix = f"_{counter}"
        candidate = f"{base[:_MAX_SHEET_TITLE_LEN - len(suffix)]}{suffix}"
        counter += 1
    return candidate


def safe_cell_value(value: Any) -> Any:
    """Return a value that openpyxl can write safely to a cell."""
    if value is None or isinstance(value, (bool, int, float, datetime, date, time, timedelta)):
        return value
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, str):
        text = ILLEGAL_CHARACTERS_RE.sub("", value)
        if text.startswith("="):
            return "'" + text
        return text
    if isinstance(value, (dict, list, tuple, set)):
        try:
            return json.dumps(value, ensure_ascii=False, sort_keys=True)
        except TypeError:
            logger.warning(
                "unsupported_export_cell_value_type",
                extra={"value_type": type(value).__name__},
            )
            return str(value)
    logger.warning(
        "unsupported_export_cell_value_type",
        extra={"value_type": type(value).__name__},
    )
    return str(value)
