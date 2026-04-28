"""Normalize router: backward-compatible alias — see standardize.py"""

from typing import Optional
from fastapi import APIRouter, Depends, Query

from webapp.dependencies import get_standardization_service
from webapp.models.responses import StandardizeResponse
from webapp.services.standardization_service import standardizationService

# This module is kept for backward compatibility.
# The canonical router is webapp/api/standardize.py
router = APIRouter(tags=["standardize"])


@router.post("/workbook/{session_id}/normalize", response_model=StandardizeResponse, include_in_schema=False)
def normalize_workbook(
    session_id: str,
    sheet: Optional[str] = Query(default=None, description="Sheet name to standardize (omit for all sheets)"),
    standardization_service: standardizationService = Depends(get_standardization_service),
) -> StandardizeResponse:
    """Backward-compatible alias for POST /standardize."""
    return standardization_service.standardize(session_id, sheet_name=sheet)
