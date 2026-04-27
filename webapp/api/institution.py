"""Institution router: GET/PATCH institution metadata and bulk MosadType apply."""

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from typing import List, Optional

from webapp.dependencies import get_session_service
from webapp.models.responses import InstitutionInfo
from webapp.services.session_service import SessionService

router = APIRouter(tags=["institution"])


class InstitutionUpdateRequest(BaseModel):
    mosad_id: Optional[str] = None
    mosad_name: Optional[str] = None
    # Up to 3 free-text user-entered institution type values.
    # The first entry is the active/default value used in export.
    mosad_types: Optional[List[str]] = None


class ApplyMosadTypeRequest(BaseModel):
    # The exact user-entered value to bulk-apply to SugMosad across all rows.
    mosad_type: str


@router.get("/workbook/{session_id}/institution", response_model=InstitutionInfo)
def get_institution(
    session_id: str,
    session_service: SessionService = Depends(get_session_service),
) -> InstitutionInfo:
    """Return the institution-level metadata for this session."""
    record = session_service.get(session_id)
    return InstitutionInfo(
        mosad_id=record.mosad_id,
        mosad_name=record.mosad_name,
        mosad_types=record.mosad_types,
    )


@router.patch("/workbook/{session_id}/institution", response_model=InstitutionInfo)
def update_institution(
    session_id: str,
    req: InstitutionUpdateRequest,
    session_service: SessionService = Depends(get_session_service),
) -> InstitutionInfo:
    """Update institution-level metadata (mosad_id, mosad_name, mosad_types)."""
    record = session_service.get(session_id)

    if req.mosad_id is not None:
        session_service.update(session_id, mosad_id=req.mosad_id.strip())
    if req.mosad_name is not None:
        session_service.update(session_id, mosad_name=req.mosad_name.strip())
    if req.mosad_types is not None:
        # Accept up to 3 values; strip whitespace; drop empty strings.
        cleaned = [v.strip() for v in req.mosad_types if v and v.strip()][:3]
        session_service.update(session_id, mosad_types=cleaned)

    record = session_service.get(session_id)
    return InstitutionInfo(
        mosad_id=record.mosad_id,
        mosad_name=record.mosad_name,
        mosad_types=record.mosad_types,
    )


@router.post("/workbook/{session_id}/mosad-type/apply")
def apply_mosad_type(
    session_id: str,
    req: ApplyMosadTypeRequest,
    session_service: SessionService = Depends(get_session_service),
) -> dict:
    """Bulk-apply a user-entered MosadType value to all SugMosad cells.

    The value must be one of the session's stored mosad_types.
    Updates every row in every sheet of the in-memory dataset so that
    the SugMosad field is set to the requested value.
    """
    record = session_service.get(session_id)

    value = req.mosad_type.strip()
    if not value:
        raise HTTPException(status_code=422, detail="mosad_type must not be empty.")

    # The applied value must be one the user has already entered.
    if value not in record.mosad_types:
        raise HTTPException(
            status_code=422,
            detail=f"'{value}' is not in the stored mosad_types for this session.",
        )

    # Bulk-update all rows in all sheets
    updated_rows = 0
    if record.workbook_dataset is not None:
        for sheet in record.workbook_dataset.sheets:
            for row in sheet.rows:
                row["SugMosad"] = value
                updated_rows += 1

    return {"mosad_type": value, "updated_rows": updated_rows}
