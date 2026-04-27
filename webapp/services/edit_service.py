"""EditService: handles manual cell edits and row deletions on in-memory SheetDataset."""

import logging
from typing import Any
from fastapi import HTTPException

from webapp.models.requests import CellEditRequest, DeleteRowRequest
from webapp.models.responses import CellEditResponse, DeleteRowResponse
from webapp.services.session_service import SessionService

logger = logging.getLogger(__name__)


def _coerce_to_original_type(new_value: str, original_value: Any) -> Any:
    """F-07: Coerce *new_value* (always a str from the API) to the type of *original_value*.

    This prevents type mismatches when editing numeric fields such as birth_year
    (originally an int) — without coercion the corrected value would be stored as
    a string, which can cause downstream export issues.

    Falls back to the raw string if coercion fails or the original type is unknown.
    """
    if isinstance(original_value, bool):
        # bool is a subclass of int; handle it first to avoid int coercion
        return new_value
    if isinstance(original_value, int):
        try:
            return int(new_value)
        except (ValueError, TypeError):
            return new_value
    if isinstance(original_value, float):
        try:
            return float(new_value)
        except (ValueError, TypeError):
            return new_value
    return new_value


class EditService:
    """Mutates in-memory SheetDataset cells and records edits in the session."""

    def __init__(self, session_service: SessionService) -> None:
        self.session_service = session_service

    def edit_cell(
        self,
        session_id: str,
        sheet_name: str,
        req: CellEditRequest,
    ) -> CellEditResponse:
        """Edit a single cell value in the in-memory dataset.

        Args:
            session_id: UUID string of the session
            sheet_name: Name of the sheet containing the cell
            req: CellEditRequest with row_uid, field_name, and new_value

        Returns:
            CellEditResponse with the updated row

        Raises:
            HTTPException 404: If session, sheet, or row_uid not found
            HTTPException 400: If field_name not found in the row
        """
        record = self.session_service.get(session_id)

        if record.workbook_dataset is None:
            raise HTTPException(
                status_code=500,
                detail="Workbook data is not available for this session.",
            )

        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)
        if sheet is None:
            raise HTTPException(
                status_code=404,
                detail=f"Sheet '{sheet_name}' not found in this workbook.",
            )

        # Find row by _row_uid
        row_idx = next(
            (i for i, r in enumerate(sheet.rows) if r.get("_row_uid") == req.row_uid),
            None,
        )
        if row_idx is None:
            raise HTTPException(
                status_code=404,
                detail=f"Row with uid '{req.row_uid}' not found in sheet '{sheet_name}'.",
            )

        # Validate field_name — must exist in the row
        row = sheet.rows[row_idx]
        if req.field_name not in row:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"Field '{req.field_name}' does not exist in sheet '{sheet_name}'. "
                    f"Available fields: {list(row.keys())}"
                ),
            )

        # F-07: Coerce new_value to the original field's type so that editing a
        # numeric field (e.g. birth_year=1990 int) stores an int, not a string.
        original_value = row.get(req.field_name)
        coerced_value: Any = _coerce_to_original_type(req.new_value, original_value)

        # Mutate the in-memory row
        sheet.rows[row_idx][req.field_name] = coerced_value

        # Record the edit in the session keyed by (sheet_name, row_uid, field_name)
        record.edits[(sheet_name, req.row_uid, req.field_name)] = coerced_value

        logger.debug(
            f"Cell edited: session={session_id}, sheet={sheet_name}, "
            f"row_uid={req.row_uid}, field={req.field_name}"
        )

        _KEEP_INTERNAL = {"_row_uid"}
        updated_row = {
            k: v for k, v in sheet.rows[row_idx].items()
            if not k.startswith("_normalization") and (not k.startswith("_") or k in _KEEP_INTERNAL)
        }
        return CellEditResponse(row_uid=req.row_uid, updated_row=updated_row)

    def delete_rows(
        self,
        session_id: str,
        sheet_name: str,
        req: DeleteRowRequest,
    ) -> DeleteRowResponse:
        """Delete one or more rows from the in-memory dataset.

        Rows are identified by their stable _row_uid strings.

        Args:
            session_id: UUID string of the session
            sheet_name: Name of the sheet to delete from
            req: DeleteRowRequest with a list of row_uids to remove

        Returns:
            DeleteRowResponse with deleted_count and remaining_rows

        Raises:
            HTTPException 404: If session or sheet not found
            HTTPException 400: If any row_uid is not found or list is empty
        """
        record = self.session_service.get(session_id)

        if record.workbook_dataset is None:
            raise HTTPException(
                status_code=500,
                detail="Workbook data is not available for this session.",
            )

        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)
        if sheet is None:
            raise HTTPException(
                status_code=404,
                detail=f"Sheet '{sheet_name}' not found in this workbook.",
            )

        if not req.row_uids:
            raise HTTPException(
                status_code=400,
                detail="row_uids must not be empty.",
            )

        uid_set = set(req.row_uids)

        # Find indices for all requested UIDs
        indices = [i for i, r in enumerate(sheet.rows) if r.get("_row_uid") in uid_set]

        # Validate all UIDs were found
        found_uids = {sheet.rows[i].get("_row_uid") for i in indices}
        missing = uid_set - found_uids
        if missing:
            raise HTTPException(
                status_code=400,
                detail=f"Row UIDs not found: {list(missing)}",
            )

        # Remove rows in reverse index order so earlier indices stay valid
        for idx in sorted(indices, reverse=True):
            sheet.rows.pop(idx)

        logger.info(
            f"Deleted {len(indices)} row(s) from sheet '{sheet_name}' "
            f"in session {session_id}. Remaining: {len(sheet.rows)}"
        )

        return DeleteRowResponse(
            deleted_count=len(indices),
            remaining_rows=len(sheet.rows),
        )
