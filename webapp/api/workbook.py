"""Workbook routers: summary, sheet data, and session management endpoints."""

from fastapi import APIRouter, Depends, Response

from webapp.dependencies import get_session_service, get_workbook_service
from webapp.models.responses import SheetDataResponse, WorkbookSummary
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
