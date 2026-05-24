"""EditService: handles manual cell edits and row deletions on in-memory SheetDataset."""

import logging
from typing import Any
from fastapi import HTTPException

from webapp.models.requests import CellEditRequest, DeleteRowRequest
from webapp.models.responses import CellEditResponse, DeleteRowResponse
from webapp.services.report_state import remove_edits_for_row_uids, sync_edit_tracking
from webapp.services.row_identity import (
    find_row_by_uid,
    missing_row_uid_error,
    row_lookup,
    row_uid,
)
from webapp.services.session_service import SessionService

logger = logging.getLogger(__name__)

_BLOCKED_FIELDS = {
    "_row_uid",
    "row_uid",
    "_validation_ok",
    "_standardization_failures",
}

_CREATABLE_GRID_FIELDS = {"MosadID", "SugMosad"}


def is_editable_source_field(field_name: str) -> bool:
    """Return whether a field can be manually edited in the Working Dataset."""
    if not field_name:
        return False
    if field_name in _BLOCKED_FIELDS:
        return False
    # Allow visible validation/status fields and corrected/status fields.
    # Block true internal prefixed fields (except _validation_status which
    # is considered user-facing when present).
    if field_name.startswith("_"):
        if field_name == "_validation_status":
            return True
        return False
    return True


# ממיר ערך ערוך מה־UI לסוג המקורי של התא ככל האפשר.
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


# שירות עריכה שמעדכן את ה־Dataset ואת עותק העבודה לפי פעולות המשתמש.
class EditService:
    """Mutates in-memory SheetDataset cells and records edits in the session."""

    # מקבל SessionService כדי לשמור שינויים על ה־session הפעיל.
    def __init__(self, session_service: SessionService) -> None:
        self.session_service = session_service

    # מעדכן תא אחד בשורת Dataset ובקובץ העבודה לפני ריצה חוזרת.
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
        logger.info(
            "cell_edit_requested",
            extra={
                "event": "cell_edit_requested",
                "session_id": session_id,
                "sheet_name": sheet_name,
                "row_uid": req.row_uid,
                "field_name": req.field_name,
            },
        )
        record = self.session_service.get(session_id)

        if record.workbook_dataset is None:
            logger.warning(
                "cell_edit_failed_no_workbook_dataset",
                extra={
                    "event": "cell_edit_failed_no_workbook_dataset",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                    "row_uid": req.row_uid,
                    "field_name": req.field_name,
                },
            )
            raise HTTPException(
                status_code=500,
                detail="Workbook data is not available for this session.",
            )

        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)
        if sheet is None:
            logger.warning(
                "cell_edit_failed_sheet_not_found",
                extra={
                    "event": "cell_edit_failed_sheet_not_found",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                    "row_uid": req.row_uid,
                    "field_name": req.field_name,
                },
            )
            raise HTTPException(
                status_code=404,
                detail=f"Sheet '{sheet_name}' not found in this workbook.",
            )

        found = find_row_by_uid(sheet, req.row_uid)
        if found is None:
            logger.warning(
                "cell_edit_failed_row_not_found",
                extra={
                    "event": "cell_edit_failed_row_not_found",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                    "row_uid": req.row_uid,
                    "field_name": req.field_name,
                },
            )
            raise missing_row_uid_error(
                sheet_name=sheet_name,
                requested_uids=[req.row_uid],
                found_count=0,
            )

        # Validate field_name — must exist in the row
        row_idx, row = found
        if req.field_name not in row and req.field_name in _CREATABLE_GRID_FIELDS:
            row[req.field_name] = ""
        if req.field_name not in row:
            logger.warning(
                "cell_edit_failed_field_not_found",
                extra={
                    "event": "cell_edit_failed_field_not_found",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                    "row_uid": req.row_uid,
                    "field_name": req.field_name,
                },
            )
            raise HTTPException(
                status_code=400,
                detail=(
                    f"Field '{req.field_name}' does not exist in sheet '{sheet_name}'. "
                    f"Available fields: {list(row.keys())}"
                ),
            )
        # Allow edits to any field present in the row (visible to the UI),
        # except for true technical/internal identifiers listed in
        # `_BLOCKED_FIELDS` or other fields explicitly considered non-editable
        # by `is_editable_source_field`.
        if not is_editable_source_field(req.field_name):
            logger.warning(
                "cell_edit_rejected_system_field",
                extra={
                    "event": "cell_edit_rejected_system_field",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                    "row_uid": req.row_uid,
                    "field_name": req.field_name,
                },
            )
            raise HTTPException(
                status_code=400,
                detail=f"Field '{req.field_name}' is an internal field and cannot be edited.",
            )

        # F-07: Coerce new_value to the original field's type so that editing a
        # numeric field (e.g. birth_year=1990 int) stores an int, not a string.
        original_value = row.get(req.field_name)
        coerced_value: Any = _coerce_to_original_type(req.new_value, original_value)

        # Mutate the in-memory row
        sheet.rows[row_idx][req.field_name] = coerced_value

        # Record the edit in the session keyed by (sheet_name, row_uid, field_name)
        sync_edit_tracking(record, sheet_name, req.row_uid, req.field_name, coerced_value)
        record.working_dataset_dirty = True
        self.session_service.update(session_id, edits=record.edits, working_dataset_dirty=True)

        logger.info(
            "cell_edit_succeeded",
            extra={
                "event": "cell_edit_succeeded",
                "session_id": session_id,
                "sheet_name": sheet_name,
                "row_uid": req.row_uid,
                "field_name": req.field_name,
                "working_dataset_dirty": True,
            },
        )

        _KEEP_INTERNAL = {"_row_uid", "_validation_status"}
        updated_row = {
            k: v for k, v in sheet.rows[row_idx].items()
            if not k.startswith("_standardization") and (not k.startswith("_") or k in _KEEP_INTERNAL)
        }
        return CellEditResponse(row_uid=req.row_uid, updated_row=updated_row)

    def update_cell(
        self,
        session_id: str,
        sheet_name: str,
        row_uid: str,
        field: str,
        value: str,
    ) -> CellEditResponse:
        """Edit one cell using the workbook-level PATCH request shape."""
        return self.edit_cell(
            session_id,
            sheet_name,
            CellEditRequest(row_uid=row_uid, field_name=field, new_value=value),
        )

    # מסמן או מסיר שורות שנמחקו כדי שלא יופיעו בתצוגה וביצוא.
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

        lookup = row_lookup(sheet)
        uid_set = set(req.row_uids)
        indices = [idx for uid, (idx, _row) in lookup.items() if uid in uid_set]

        # Validate all UIDs were found
        found_uids = {row_uid(sheet.rows[i]) for i in indices}
        missing = uid_set - found_uids
        if missing:
            raise missing_row_uid_error(
                sheet_name=sheet_name,
                requested_uids=req.row_uids,
                found_count=len(found_uids),
                status_code=400,
            )

        # Remove rows in reverse index order so earlier indices stay valid
        for idx in sorted(indices, reverse=True):
            sheet.rows.pop(idx)
        remove_edits_for_row_uids(record, sheet_name, req.row_uids)
        record.working_dataset_dirty = True
        self.session_service.update(
            session_id,
            edits=record.edits,
            working_dataset_dirty=True,
        )

        logger.info(
            "rows_deleted",
            extra={
                "event": "rows_deleted",
                "session_id": session_id,
                "sheet_name": sheet_name,
                "deleted_count": len(indices),
                "remaining_rows": len(sheet.rows),
                "working_dataset_dirty": True,
            },
        )

        return DeleteRowResponse(
            deleted_count=len(indices),
            remaining_rows=len(sheet.rows),
        )
