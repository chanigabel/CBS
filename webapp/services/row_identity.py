"""Stable row identity helpers shared by grid/edit/export flows."""

from __future__ import annotations

import uuid
from typing import Any, Iterable

from fastapi import HTTPException

ROW_UID_FIELD = "_row_uid"
LEGACY_ROW_UID_FIELD = "row_uid"


def _clean_uid(value: Any) -> str:
    return str(value or "").strip()


def ensure_row_uid(row: dict[str, Any], used_uids: set[str] | None = None) -> str:
    """Ensure a row has one stable backend-owned UID.

    ``_row_uid`` is the canonical internal field. ``row_uid`` is accepted as a
    legacy alias for lookup/import compatibility, but new IDs are not written
    to ``row_uid`` so it does not become a visible source column.
    """
    used_uids = used_uids if used_uids is not None else set()
    uid = _clean_uid(row.get(ROW_UID_FIELD)) or _clean_uid(row.get(LEGACY_ROW_UID_FIELD))
    if not uid or uid in used_uids:
        uid = uuid.uuid4().hex
    row[ROW_UID_FIELD] = uid
    used_uids.add(uid)
    return uid


def ensure_sheet_row_uids(sheet) -> None:
    """Create and persist missing row UIDs on the session SheetDataset."""
    used_uids: set[str] = set()
    for row in sheet.rows:
        ensure_row_uid(row, used_uids)


def row_uid(row: dict[str, Any]) -> str:
    """Return the canonical or legacy UID for a row without creating one."""
    return _clean_uid(row.get(ROW_UID_FIELD)) or _clean_uid(row.get(LEGACY_ROW_UID_FIELD))


def row_lookup(sheet) -> dict[str, tuple[int, dict[str, Any]]]:
    """Return a UID -> (index, row) lookup for a sheet, creating missing UIDs."""
    ensure_sheet_row_uids(sheet)
    lookup: dict[str, tuple[int, dict[str, Any]]] = {}
    for index, row in enumerate(sheet.rows):
        uid = row_uid(row)
        if uid:
            lookup[uid] = (index, row)
    return lookup


def find_row_by_uid(sheet, uid: str) -> tuple[int, dict[str, Any]] | None:
    """Find a row by canonical/legacy UID in the same data backing the grid."""
    return row_lookup(sheet).get(_clean_uid(uid))


def missing_row_uid_error(
    *,
    sheet_name: str,
    requested_uids: Iterable[str],
    found_count: int,
    status_code: int = 404,
) -> HTTPException:
    requested = list(requested_uids)
    missing_count = max(len(requested) - found_count, 0)
    return HTTPException(
        status_code=status_code,
        detail={
            "message": "הבחירה בגריד אינה מעודכנת. נא לבחור את השורות מחדש ולנסות שוב.",
            "sheet_name": sheet_name,
            "requested_rows": len(requested),
            "found_rows": found_count,
            "missing_rows": missing_count,
        },
    )
