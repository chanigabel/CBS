"""Institution router: GET/PATCH institution metadata and bulk MosadType apply."""

import re
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, field_validator
from typing import List, Optional

from webapp.dependencies import get_session_service
from webapp.models.requests import ApplySugMosadRequest, _validate_numeric_min3
from webapp.models.responses import InstitutionInfo
from webapp.models.session import SugMosadConfig, SelectedRowsConfig
from webapp.services.session_service import SessionService

router = APIRouter(tags=["institution"])


class InstitutionUpdateRequest(BaseModel):
    mosad_id: Optional[str] = None
    mosad_name: Optional[str] = None
    mosad_types: Optional[List[str]] = None

    @field_validator("mosad_id")
    @classmethod
    def validate_mosad_id(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return v
        return _validate_numeric_min3(v, "מספר מוסד")

    @field_validator("mosad_types")
    @classmethod
    def validate_mosad_types(cls, v: Optional[List[str]]) -> Optional[List[str]]:
        if v is None:
            return v
        validated = []
        for item in v:
            validated.append(_validate_numeric_min3(item, "סוג מוסד"))
        return validated


class ApplyMosadTypeRequest(BaseModel):
    mosad_type: str

    @field_validator("mosad_type")
    @classmethod
    def validate_mosad_type(cls, v: str) -> str:
        return _validate_numeric_min3(v, "סוג מוסד")


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
    """Bulk-apply a user-entered MosadType value to all SugMosad cells (legacy)."""
    record = session_service.get(session_id)

    value = req.mosad_type.strip()
    if not value:
        raise HTTPException(status_code=422, detail="mosad_type must not be empty.")

    if value not in record.mosad_types:
        raise HTTPException(
            status_code=422,
            detail=f"'{value}' is not in the stored mosad_types for this session.",
        )

    updated_rows = 0
    if record.workbook_dataset is not None:
        for sheet in record.workbook_dataset.sheets:
            for row in sheet.rows:
                row["SugMosad"] = value
                updated_rows += 1

    return {"mosad_type": value, "updated_rows": updated_rows}


@router.post("/workbook/{session_id}/mosad-type/apply-scoped")
def apply_mosad_type_scoped(
    session_id: str,
    req: ApplySugMosadRequest,
    session_service: SessionService = Depends(get_session_service),
) -> dict:
    """Apply SugMosad with a specific scope: workbook, sheet, or selected rows.

    Scope "workbook":
        Applies sug_mosad to every row in every sheet.

    Scope "sheet":
        Applies sug_mosad only to rows in sheet_name.
        Other sheets are untouched.

    Scope "selected_rows":
        Applies each entry's sug_mosad to the rows matching its row_uids
        inside sheet_name.  Uses the same _row_uid identity as row deletion.
        Other sheets and unselected rows are untouched.
        Up to 3 SelectedRowsRequest entries allowed.

    In all cases the config is persisted on the session so export reproduces
    the exact same scoped application.
    """
    record = session_service.get(session_id)

    if record.workbook_dataset is None:
        raise HTTPException(
            status_code=422,
            detail="Workbook data not loaded. Please load a sheet first.",
        )

    if req.mosad_id:
        session_service.update(session_id, mosad_id=req.mosad_id.strip())

    updated_rows = 0

    if req.scope == "workbook":
        for sheet in record.workbook_dataset.sheets:
            for row in sheet.rows:
                row["SugMosad"] = req.sug_mosad
                updated_rows += 1
        config = SugMosadConfig(scope="workbook", sug_mosad=req.sug_mosad)

    elif req.scope == "sheet":
        sheet_obj = record.workbook_dataset.get_sheet_by_name(req.sheet_name)
        if sheet_obj is None:
            raise HTTPException(status_code=404, detail=f"Sheet '{req.sheet_name}' not found.")
        for row in sheet_obj.rows:
            row["SugMosad"] = req.sug_mosad
            updated_rows += 1
        config = SugMosadConfig(
            scope="sheet",
            sug_mosad=req.sug_mosad,
            sheet_name=req.sheet_name,
        )

    else:  # scope == "selected_rows"
        sheet_obj = record.workbook_dataset.get_sheet_by_name(req.sheet_name)
        if sheet_obj is None:
            raise HTTPException(status_code=404, detail=f"Sheet '{req.sheet_name}' not found.")

        # Build a uid->row lookup for fast access (same identity as row deletion)
        uid_to_row = {
            row["_row_uid"]: row
            for row in sheet_obj.rows
            if "_row_uid" in row
        }

        selected_configs: List[SelectedRowsConfig] = []
        for entry in req.selected_rows:
            uid_set = set(entry.row_uids)
            matched = 0
            for uid in entry.row_uids:
                row = uid_to_row.get(uid)
                if row is not None:
                    row["SugMosad"] = entry.sug_mosad
                    matched += 1
            updated_rows += matched
            selected_configs.append(
                SelectedRowsConfig(sug_mosad=entry.sug_mosad, row_uids=list(uid_set))
            )

        # Merge into any existing selected_rows config for this sheet so that
        # repeated calls (select rows → apply 123, select other rows → apply 456)
        # accumulate rather than overwrite.  New entries for the same sug_mosad
        # value extend the existing uid list; new sug_mosad values are appended.
        existing_configs = record.sug_mosad_configs
        existing_sr = next(
            (c for c in existing_configs
             if c.scope == "selected_rows" and c.sheet_name == req.sheet_name),
            None,
        )
        if existing_sr is not None:
            # Build a map of sug_mosad → set of row_uids from the existing config
            existing_map: dict = {}
            for grp in existing_sr.selected_rows:
                existing_map.setdefault(grp.sug_mosad, set()).update(grp.row_uids)
            # Merge new entries: new UIDs override any previous sug_mosad assignment
            # for those UIDs (a row can only have one sug_mosad at a time).
            new_uids_all: set = set()
            for sc in selected_configs:
                new_uids_all.update(sc.row_uids)
            # Remove newly-assigned UIDs from all existing groups
            for key in list(existing_map.keys()):
                existing_map[key] -= new_uids_all
                if not existing_map[key]:
                    del existing_map[key]
            # Add new groups
            for sc in selected_configs:
                existing_map.setdefault(sc.sug_mosad, set()).update(sc.row_uids)
            # Rebuild selected_configs from merged map (cap at 3 groups)
            merged = [
                SelectedRowsConfig(sug_mosad=sm, row_uids=list(uids))
                for sm, uids in existing_map.items()
                if uids
            ][:3]
            config = SugMosadConfig(
                scope="selected_rows",
                sheet_name=req.sheet_name,
                selected_rows=merged,
            )
        else:
            config = SugMosadConfig(
                scope="selected_rows",
                sheet_name=req.sheet_name,
                selected_rows=selected_configs,
            )

    # Persist — replace any existing config for the same scope/sheet.
    existing = record.sug_mosad_configs
    new_configs = [
        c for c in existing
        if not (c.scope == config.scope and c.sheet_name == config.sheet_name)
    ]
    new_configs.append(config)
    session_service.update(session_id, sug_mosad_configs=new_configs)

    return {
        "scope": req.scope,
        "sheet_name": req.sheet_name,
        "updated_rows": updated_rows,
    }

