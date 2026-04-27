"""Edit router: cell edits and row deletions."""

from fastapi import APIRouter, Depends

from webapp.dependencies import get_edit_service
from webapp.models.requests import CellEditRequest, DeleteRowRequest
from webapp.models.responses import CellEditResponse, DeleteRowResponse
from webapp.services.edit_service import EditService

router = APIRouter(tags=["edit"])


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
