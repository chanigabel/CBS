"""Service for multi-row editing operations in the grid."""

from __future__ import annotations

import logging
from typing import List, Optional, Dict, Any, Tuple
from fastapi import HTTPException

from webapp.models.requests import CellEditRequest
from webapp.services.edit_service import is_editable_source_field, _coerce_to_original_type
from webapp.services.report_state import sync_edit_tracking
from webapp.services.row_identity import missing_row_uid_error, row_lookup
from webapp.services.session_service import SessionService

logger = logging.getLogger(__name__)
_CREATABLE_GRID_FIELDS = {"MosadID", "SugMosad"}


class MultiRowEditRequest:
    """Request to edit the same cell in multiple rows."""
    
    def __init__(
        self,
        row_uids: List[str],
        field_name: str,
        new_value: str,
    ):
        """Initialize a multi-row edit request.
        
        Args:
            row_uids: List of row UIDs to edit
            field_name: Name of the field to edit
            new_value: New value to apply to all rows
        """
        self.row_uids = row_uids
        self.field_name = field_name
        self.new_value = new_value


class MultiRowEditResult:
    """Result of a multi-row edit operation."""
    
    def __init__(self, edited_count: int, updated_rows: Dict[str, Dict[str, Any]]):
        """Initialize a multi-row edit result.
        
        Args:
            edited_count: Number of rows successfully edited
            updated_rows: Dict mapping row_uid to updated row data
        """
        self.edited_count = edited_count
        self.updated_rows = updated_rows


class MultiRowEditService:
    """Service for editing multiple rows at once."""

    def __init__(self, session_service: SessionService) -> None:
        self.session_service = session_service

    def edit_multiple_rows(
        self,
        session_id: str,
        sheet_name: str,
        row_uids: List[str],
        field_name: str,
        new_value: str,
    ) -> MultiRowEditResult:
        """Edit the same field in multiple rows.

        Args:
            session_id: UUID string of the session
            sheet_name: Name of the sheet containing the rows
            row_uids: List of row UIDs to edit
            field_name: Name of the field to edit in each row
            new_value: New value to apply to all rows

        Returns:
            MultiRowEditResult with count and updated rows

        Raises:
            HTTPException: If validation fails or data not available
        """
        logger.info(
            "multi_row_edit_requested",
            extra={
                "event": "multi_row_edit_requested",
                "session_id": session_id,
                "sheet_name": sheet_name,
                "row_count": len(row_uids),
                "field_name": field_name,
            },
        )

        record = self.session_service.get(session_id)

        if record.workbook_dataset is None:
            logger.warning(
                "multi_row_edit_failed_no_workbook_dataset",
                extra={
                    "event": "multi_row_edit_failed_no_workbook_dataset",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                },
            )
            raise HTTPException(
                status_code=500,
                detail="Workbook data is not available for this session.",
            )

        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)
        if sheet is None:
            logger.warning(
                "multi_row_edit_failed_sheet_not_found",
                extra={
                    "event": "multi_row_edit_failed_sheet_not_found",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                },
            )
            raise HTTPException(
                status_code=404,
                detail=f"הגיליון '{sheet_name}' לא נמצא בקובץ.",
            )

        # Validate field_name is editable
        if not is_editable_source_field(field_name):
            logger.warning(
                "multi_row_edit_rejected_system_field",
                extra={
                    "event": "multi_row_edit_rejected_system_field",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                    "field_name": field_name,
                },
            )
            raise HTTPException(
                status_code=400,
                detail=f"Field '{field_name}' is an internal field and cannot be edited.",
            )

        uid_to_row = row_lookup(sheet)

        # Validate all row_uids exist
        missing_uids = [uid for uid in row_uids if uid not in uid_to_row]
        if missing_uids:
            logger.warning(
                "multi_row_edit_failed_rows_not_found",
                extra={
                    "event": "multi_row_edit_failed_rows_not_found",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                    "missing_uids": missing_uids,
                },
            )
            raise missing_row_uid_error(
                sheet_name=sheet_name,
                requested_uids=row_uids,
                found_count=len(row_uids) - len(missing_uids),
            )

        # Validate field_name exists in all rows
        for uid in row_uids:
            _idx, row = uid_to_row[uid]
            if field_name not in row and field_name in _CREATABLE_GRID_FIELDS:
                row[field_name] = ""
            if field_name not in row:
                logger.warning(
                    "multi_row_edit_failed_field_not_found",
                    extra={
                        "event": "multi_row_edit_failed_field_not_found",
                        "session_id": session_id,
                        "sheet_name": sheet_name,
                        "row_uid": uid,
                        "field_name": field_name,
                    },
                )
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"Field '{field_name}' does not exist in one or more rows. "
                        f"Cannot apply edit to {len(row_uids)} selected rows."
                    ),
                )

        # Apply the edit to all rows
        updated_rows: Dict[str, Dict[str, Any]] = {}
        for uid in row_uids:
            _idx, row = uid_to_row[uid]

            # Coerce new value to match original type
            original_value = row.get(field_name)
            coerced_value = _coerce_to_original_type(new_value, original_value)

            # Update the row
            row[field_name] = coerced_value

            # Track edit in session edits dict: key is (sheet_name, row_uid, field_name)
            edit_key = (sheet_name, uid, field_name)
            sync_edit_tracking(record, sheet_name, uid, field_name, coerced_value)

            # Collect updated row for response
            updated_rows[uid] = dict(row)

        # Mark working dataset as dirty
        record.working_dataset_dirty = True
        self.session_service.update(
            session_id,
            edits=record.edits,
            working_dataset_dirty=True,
        )

        logger.info(
            "multi_row_edit_completed",
            extra={
                "event": "multi_row_edit_completed",
                "session_id": session_id,
                "sheet_name": sheet_name,
                "rows_edited": len(row_uids),
                "field_name": field_name,
            },
        )

        return MultiRowEditResult(
            edited_count=len(row_uids),
            updated_rows=updated_rows,
        )


class UndoStack:
    """Track undo/redo state for grid edits."""

    def __init__(self):
        """Initialize an empty undo stack."""
        self._undo_stack: List[Tuple[str, str, List[str], str, Dict[str, Any]]] = []
        self._redo_stack: List[Tuple[str, str, List[str], str, Dict[str, Any]]] = []

    def push_edit(
        self,
        sheet_name: str,
        row_uids: List[str],
        field_name: str,
        new_value: str,
        old_values: Dict[str, Any],
    ) -> None:
        """Push an edit action onto the undo stack.

        Args:
            sheet_name: Sheet name
            row_uids: List of affected row UIDs
            field_name: Field that was edited
            new_value: The new value
            old_values: Dict mapping row_uid to old value
        """
        self._undo_stack.append((sheet_name, new_value, row_uids, field_name, old_values))
        self._redo_stack.clear()  # Clear redo when new edit happens

    def can_undo(self) -> bool:
        """Return True if there are undoable edits."""
        return len(self._undo_stack) > 0

    def undo(self) -> Optional[Tuple[str, str, List[str], str, Dict[str, Any]]]:
        """Undo the last edit.

        Returns:
            Tuple of (sheet_name, old_value, row_uids, field_name, old_values)
            or None if nothing to undo
        """
        if not self.can_undo():
            return None

        action = self._undo_stack.pop()
        sheet_name, new_value, row_uids, field_name, old_values = action
        self._redo_stack.append(action)

        logger.info(
            "undo_performed",
            extra={
                "event": "undo_performed",
                "sheet_name": sheet_name,
                "field_name": field_name,
                "rows_affected": len(row_uids),
            },
        )

        return (sheet_name, new_value, row_uids, field_name, old_values)

    def can_redo(self) -> bool:
        """Return True if there are redoable edits."""
        return len(self._redo_stack) > 0

    def redo(self) -> Optional[Tuple[str, str, List[str], str, Dict[str, Any]]]:
        """Redo the last undone edit.

        Returns:
            Tuple of (sheet_name, new_value, row_uids, field_name, old_values)
            or None if nothing to redo
        """
        if not self.can_redo():
            return None

        action = self._redo_stack.pop()
        sheet_name, new_value, row_uids, field_name, old_values = action
        self._undo_stack.append(action)

        logger.info(
            "redo_performed",
            extra={
                "event": "redo_performed",
                "sheet_name": sheet_name,
                "field_name": field_name,
                "rows_affected": len(row_uids),
            },
        )

        return action
