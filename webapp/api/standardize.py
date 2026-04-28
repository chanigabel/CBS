"""Standardize router: POST /api/workbook/{session_id}/standardize"""

from typing import Optional
from fastapi import APIRouter, Depends, Query

from webapp.dependencies import get_standardization_service
from webapp.models.responses import StandardizeResponse
from webapp.services.standardization_service import standardizationService

router = APIRouter(tags=["standardize"])


@router.post("/workbook/{session_id}/standardize", response_model=StandardizeResponse)
def standardize_workbook(
    session_id: str,
    sheet: Optional[str] = Query(default=None, description="Sheet name to standardize (omit for all sheets)"),
    standardization_service: standardizationService = Depends(get_standardization_service),
) -> StandardizeResponse:
    """Run the standardization pipeline on the session's working copy.

    Pass ?sheet=<name> to standardize only the active sheet (faster).
    Omit the parameter to standardize all sheets.
    """
    return standardization_service.standardize(session_id, sheet_name=sheet)


# Backward-compatible alias — keeps existing frontend/clients working
@router.post("/workbook/{session_id}/normalize", response_model=StandardizeResponse, include_in_schema=False)
def normalize_workbook_alias(
    session_id: str,
    sheet: Optional[str] = Query(default=None),
    standardization_service: standardizationService = Depends(get_standardization_service),
) -> StandardizeResponse:
    """Backward-compatible alias for /standardize."""
    return standardization_service.standardize(session_id, sheet_name=sheet)
