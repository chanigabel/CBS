"""Field matching helpers for ExcelReader."""

from __future__ import annotations

import re
from typing import Any, Callable, Iterable, Optional

from openpyxl.worksheet.worksheet import Worksheet


def normalize_text(text: str) -> str:
    """Normalize text for comparison."""
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"[()[\]{}]", " ", text)
    text = text.lower()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def contains_field_keyword(
    normalized_text: str,
    field_keywords: dict[str, list[str]],
) -> bool:
    """Check if normalized text contains any field keyword."""
    for keywords in field_keywords.values():
        for keyword in keywords:
            if keyword in normalized_text:
                return True
    return False


def should_ignore_column(cell_text: str, ignore_keywords: Iterable[str]) -> bool:
    """Check if a column should be ignored."""
    normalized = normalize_text(cell_text)
    for ignore_word in ignore_keywords:
        if ignore_word in normalized:
            return True
    return False


def looks_like_data_value(cell_value: Any) -> bool:
    """Return True if a cell value looks like a data value rather than a header."""
    from datetime import datetime as _dt, date as _date

    if isinstance(cell_value, (_dt, _date)):
        return True
    if isinstance(cell_value, (int, float)):
        return True
    if cell_value is None:
        return False

    txt = str(cell_value).strip()
    if not txt:
        return False

    import re as _re
    if _re.match(r"^\d{4}-\d{2}-\d{2}", txt):
        return True
    if _re.match(r"^\d{1,2}[./]\d{1,2}[./]\d{2,4}$", txt):
        return True
    if txt.isdigit():
        return True

    stripped = txt.replace(",", "").replace(".", "").replace("-", "")
    if stripped.isdigit() and len(stripped) > 0:
        return True

    return False


def find_label_row(
    worksheet: Worksheet,
    col: int,
    header_area_rows: list,
    looks_like_data_value_fn: Callable[[Any], bool],
) -> int:
    """Return the row in header_area_rows where col has its label."""
    for hr in header_area_rows:
        v = worksheet.cell(row=hr, column=col).value
        if v is not None and str(v).strip() != "" and not looks_like_data_value_fn(v):
            return hr
    return header_area_rows[0] if header_area_rows else 1
