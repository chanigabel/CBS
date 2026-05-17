"""Workbook routers: summary, sheet data, and session management endpoints."""

from fastapi import APIRouter, Depends, Response

from webapp.dependencies import get_edit_service, get_session_service, get_workbook_service
from webapp.models.requests import ColumnMappingRequest, WorkbookCellUpdateRequest
from webapp.models.responses import (
    CellEditResponse,
    ColumnMappingResponse,
    ColumnSchemaResponse,
    SheetDataResponse,
    WorkbookSummary,
)
from webapp.services.edit_service import EditService
from webapp.services.session_service import SessionService
from webapp.services.workbook_service import WorkbookService

router = APIRouter(tags=["workbook"])


@router.get("/workbook/{session_id}/summary", response_model=WorkbookSummary)
def get_workbook_summary(
    session_id: str,
    workbook_service: WorkbookService = Depends(get_workbook_service),
) -> WorkbookSummary:
    """Return a summary of all sheets in the uploaded workbook."""
    return workbook_service.get_summary(session_id)


@router.get(
    "/workbook/{session_id}/sheet/{sheet_name}",
    response_model=SheetDataResponse,
)
def get_sheet_data(
    session_id: str,
    sheet_name: str,
    workbook_service: WorkbookService = Depends(get_workbook_service),
) -> SheetDataResponse:
    """Return all rows for a specific sheet."""
    return workbook_service.get_sheet_data(session_id, sheet_name)


@router.patch(
    "/workbook/{session_id}/cell",
    response_model=CellEditResponse,
)
def update_workbook_cell(
    session_id: str,
    request: WorkbookCellUpdateRequest,
    edit_service: EditService = Depends(get_edit_service),
) -> CellEditResponse:
    """Update one editable source cell in the session Working Dataset."""
    return edit_service.update_cell(
        session_id,
        request.sheet_name,
        request.row_uid,
        request.field,
        request.value,
    )


@router.get("/workbook/column-schema", response_model=ColumnSchemaResponse)
def get_column_schema(
    workbook_service: WorkbookService = Depends(get_workbook_service),
) -> ColumnSchemaResponse:
    """Return standardized field names available for manual column mapping."""
    return workbook_service.get_column_schema()


@router.post(
    "/workbook/{session_id}/sheet/{sheet_name}/column-mapping",
    response_model=ColumnMappingResponse,
)
def update_column_mapping(
    session_id: str,
    sheet_name: str,
    request: ColumnMappingRequest,
    workbook_service: WorkbookService = Depends(get_workbook_service),
) -> ColumnMappingResponse:
    """Map one loaded source column to a standardized field name."""
    return workbook_service.update_column_mapping(
        session_id,
        sheet_name,
        request.old_name,
        request.new_name,
    )


@router.post(
    "/workbook/{session_id}/sheet/{sheet_name}/reload-mapping",
    response_model=ColumnSchemaResponse,
)
def reload_column_mapping(
    session_id: str,
    sheet_name: str,
    workbook_service: WorkbookService = Depends(get_workbook_service),
) -> ColumnSchemaResponse:
    """Reload column mapping schema and re-apply stored mappings to this sheet."""
    return workbook_service.reload_column_mapping(session_id, sheet_name)


@router.delete("/workbook/{session_id}", status_code=204)
def close_session(
    session_id: str,
    session_service: SessionService = Depends(get_session_service),
) -> Response:
    """F-08: Remove a session from memory.

    Frees the in-memory WorkbookDataset for this session.
    Does NOT delete the uploaded source or working-copy files from disk.
    Returns 204 No Content on success (including when the session does not exist,
    to make the operation idempotent from the client's perspective).
    """
    session_service.delete(session_id)
    return Response(status_code=204)
