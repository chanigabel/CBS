"""Pydantic response models for the API layer."""

from typing import Any, Dict, List
from pydantic import BaseModel


class UploadResponse(BaseModel):
    """Response from POST /api/upload."""

    session_id: str
    sheet_names: List[str]


class SheetSummary(BaseModel):
    """Summary of a single sheet within a workbook."""

    sheet_name: str
    row_count: int
    field_names: List[str]


class WorkbookSummary(BaseModel):
    """Response from GET /api/workbook/{session_id}/summary."""

    session_id: str
    sheets: List[SheetSummary]


class SheetDataResponse(BaseModel):
    """Response from GET /api/workbook/{session_id}/sheet/{sheet_name}."""

    sheet_name: str
    field_names: List[str]
    rows: List[Dict[str, Any]]


class PerSheetStat(BaseModel):
    """Per-sheet normalization statistics."""

    sheet_name: str
    rows: int
    success_rate: float


class NormalizeResponse(BaseModel):
    """Response from POST /api/workbook/{session_id}/normalize."""

    session_id: str
    status: str
    sheets_processed: int
    total_rows: int
    per_sheet_stats: List[PerSheetStat]


class CellEditResponse(BaseModel):
    """Response from PATCH /api/workbook/{session_id}/sheet/{sheet_name}/cell."""

    row_uid: str
    updated_row: Dict[str, Any]


class DeleteRowResponse(BaseModel):
    """Response from DELETE /api/workbook/{session_id}/sheet/{sheet_name}/rows."""

    deleted_count: int
    remaining_rows: int


class InstitutionInfo(BaseModel):
    """Institution-level metadata for a workbook session."""

    mosad_id: str = ""
    mosad_name: str = ""
    mosad_types: List[str] = []


class ErrorResponse(BaseModel):
    """Standard error response body."""

    detail: str
