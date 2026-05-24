"""Edit router: cell edits, multi-row edits, and row deletions."""

from typing import List
from pydantic import BaseModel
from fastapi import APIRouter, Depends

from webapp.dependencies import get_edit_service, get_multirow_edit_service
from webapp.models.requests import CellEditRequest, DeleteRowRequest
from webapp.models.responses import CellEditResponse, DeleteRowResponse
from webapp.services.edit_service import EditService
from webapp.services.multirow_edit_service import MultiRowEditService

router = APIRouter(tags=["edit"])


class MultiRowEditRequest(BaseModel):
    """Request to edit multiple rows."""
    row_uids: List[str]
    field_name: str
    new_value: str


class MultiRowEditResponse(BaseModel):
    """Response from multi-row edit."""
    edited_count: int
    sheet_name: str
    field_name: str


@router.patch(
    "/workbook/{session_id}/sheet/{sheet_name}/cell",
    response_model=CellEditResponse,
)
def edit_cell(
    session_id: str,
    sheet_name: str,
    req: CellEditRequest,
    edit_service: EditService = Depends(get_edit_service),
) -> CellEditResponse:
    """Edit a single cell value in the in-memory dataset."""
    return edit_service.edit_cell(session_id, sheet_name, req)


@router.patch(
    "/workbook/{session_id}/sheet/{sheet_name}/multi-edit",
    response_model=MultiRowEditResponse,
)
def edit_multiple_rows(
    session_id: str,
    sheet_name: str,
    req: MultiRowEditRequest,
    multirow_service: MultiRowEditService = Depends(get_multirow_edit_service),
) -> MultiRowEditResponse:
    """Edit the same field in multiple selected rows at once.
    
    Args:
        session_id: Session UUID
        sheet_name: Sheet name
        req: MultiRowEditRequest with row_uids, field_name, and new_value
        
    Returns:
        MultiRowEditResponse with count of edited rows
    """
    result = multirow_service.edit_multiple_rows(
        session_id, sheet_name, req.row_uids, req.field_name, req.new_value
    )
    return MultiRowEditResponse(
        edited_count=result.edited_count,
        sheet_name=sheet_name,
        field_name=req.field_name,
    )


@router.delete(
    "/workbook/{session_id}/sheet/{sheet_name}/rows",
    response_model=DeleteRowResponse,
)
def delete_rows(
    session_id: str,
    sheet_name: str,
    req: DeleteRowRequest,
    edit_service: EditService = Depends(get_edit_service),
) -> DeleteRowResponse:
    """Delete one or more rows from the in-memory dataset.

    Pass a JSON body with ``row_uids``: a list of stable row UID strings
    to delete.  All UIDs are validated before any deletion occurs.
    """
    return edit_service.delete_rows(session_id, sheet_name, req)
