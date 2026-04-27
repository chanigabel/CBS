"""Normalize router: POST /api/workbook/{session_id}/normalize"""

from typing import Optional
from fastapi import APIRouter, Depends, Query

from webapp.dependencies import get_normalization_service
from webapp.models.responses import NormalizeResponse
from webapp.services.normalization_service import NormalizationService

router = APIRouter(tags=["normalize"])


@router.post("/workbook/{session_id}/normalize", response_model=NormalizeResponse)
def normalize_workbook(
    session_id: str,
    sheet: Optional[str] = Query(default=None, description="Sheet name to normalize (omit for all sheets)"),
    normalization_service: NormalizationService = Depends(get_normalization_service),
) -> NormalizeResponse:
    """Run the normalization pipeline on the session's working copy.

    Pass ?sheet=<name> to normalize only the active sheet (faster).
    Omit the parameter to normalize all sheets.
    """
    return normalization_service.normalize(session_id, sheet_name=sheet)
