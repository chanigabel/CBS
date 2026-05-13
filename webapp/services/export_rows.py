"""Helpers for export filename generation and row preparation."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from webapp.services.derived_columns import apply_derived_columns, detect_serial_field, SYNTHETIC_SERIAL_KEY
from webapp.services.export_schema import canonical_sheet_name, headers_for_sheet
from webapp.services.export_validation import (
    is_numeric_like,
    row_has_visible_original_values,
    row_is_numeric_helper_row,
)


# ממיר שם בסיס לשם שדה יצוא בפורמט PascalCase.
def _to_pascal_case(text: str) -> str:
    """Convert a free-text name to PascalCase English words joined without spaces."""
    import re

    tokens = re.split(r"[\s\-]+", text.strip())
    return "".join(t.capitalize() for t in tokens if t)


def _safe_filename_stem(text: str, fallback: str = "export") -> str:
    """Keep Unicode names but remove characters invalid in Windows filenames."""
    invalid = '<>:"/\\|?*'
    cleaned = ILLEGAL_CHARACTERS_RE.sub("", text)
    cleaned = "".join("_" if ch in invalid else ch for ch in cleaned)
    cleaned = " ".join(cleaned.split()).strip(" ._")
    return cleaned or fallback


# בונה שם קובץ יצוא ייחודי לפי שם המקור וזמן הריצה.
def build_export_filename(record) -> str:
    """Build the export filename from institution metadata."""
    mosad_id = _safe_filename_stem(record.mosad_id or "", fallback="").strip()
    mosad_name = _safe_filename_stem(record.mosad_name or "", fallback="").strip()

    if mosad_id and mosad_name:
        pascal = _to_pascal_case(mosad_name)
        return _safe_filename_stem(f"{mosad_id} {pascal}") + ".xlsx"

    original_stem = _safe_filename_stem(Path(record.original_filename).stem)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{original_stem}_standardized_{timestamp}.xlsx"


# קובע SugMosad לגיליון לפי קונפיגורציית היצוא או fallback.
def resolve_sug_mosad_for_sheet(configs, sheet_name: str, fallback: str):
    """Return the SugMosad value (or callable) to apply for a given sheet during export."""
    if not configs:
        return fallback

    for cfg in configs:
        if cfg.scope == "selected_rows" and cfg.sheet_name == sheet_name:
            uid_map: dict = {}
            for grp in cfg.selected_rows:
                for uid in grp.row_uids:
                    uid_map[uid] = grp.sug_mosad

            def _uid_resolver(row_uid: str, _map=uid_map) -> Optional[str]:
                return _map.get(row_uid)

            return _uid_resolver

    for cfg in configs:
        if cfg.scope == "sheet" and cfg.sheet_name == sheet_name:
            return cfg.sug_mosad

    for cfg in configs:
        if cfg.scope == "workbook":
            return cfg.sug_mosad

    return fallback


def build_row_export_view(row: Dict[str, Any], mosad_id: str = "", scoped_sug_mosad=None) -> Dict[str, Any]:
    """Return a non-mutating row view with export-scoped institution metadata."""
    export_row = dict(row)
    if mosad_id:
        export_row["MosadID"] = mosad_id
    if callable(scoped_sug_mosad):
        value = scoped_sug_mosad(row.get("_row_uid", ""))
        if value is not None:
            export_row["SugMosad"] = value
    elif scoped_sug_mosad:
        export_row["SugMosad"] = scoped_sug_mosad
    return export_row


# מחזיר רק שורות פעילות ושדות מקור שרלוונטיים ליצוא ולדוחות.
def visible_rows(sheet_dataset) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Return (rows, display_columns) exactly as the UI would show them."""
    original_field_set = set(sheet_dataset.field_names)

    rows = [
        {k: v for k, v in row.items() if not k.startswith("_standardization")}
        for row in sheet_dataset.rows
    ]

    rows = [
        row for row in rows
        if row_has_visible_original_values(row, original_field_set)
    ]

    if rows and row_is_numeric_helper_row(rows[0], original_field_set):
        rows = rows[1:]

    display_columns = list(sheet_dataset.field_names)
    meta_mosad_id = sheet_dataset.get_metadata("MosadID")
    rows, display_columns = apply_derived_columns(
        rows=rows,
        field_names=sheet_dataset.field_names,
        display_columns=display_columns,
        meta_mosad_id=meta_mosad_id,
    )
    return rows, display_columns

