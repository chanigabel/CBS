"""Pydantic request models for the API layer."""

import re
from typing import List, Optional
from pydantic import BaseModel, field_validator, model_validator


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------

def _validate_numeric_min3(value: str, field_label: str) -> str:
    """Raise ValueError if value is not numeric-only with at least 3 digits."""
    stripped = value.strip()
    if not stripped:
        return stripped
    if not re.fullmatch(r'\d+', stripped):
        raise ValueError(f"{field_label} חייב להכיל ספרות בלבד")
    if len(stripped) < 3:
        raise ValueError(f"{field_label} חייב להכיל לפחות 3 ספרות")
    return stripped


# ---------------------------------------------------------------------------
# Cell / row edit models
# ---------------------------------------------------------------------------

class CellEditRequest(BaseModel):
    """Request body for editing a single cell value."""
    row_uid: str
    field_name: str
    new_value: str


class WorkbookCellUpdateRequest(BaseModel):
    """Request body for editing one Working Dataset cell."""

    sheet_name: str
    row_uid: str
    field: str
    value: str


class DeleteRowRequest(BaseModel):
    """Request body for deleting one or more rows from a sheet."""
    row_uids: List[str]


class ColumnMappingRequest(BaseModel):
    """Request body for renaming one source column to a standardized field."""

    old_name: str
    new_name: str


class ColumnSchemaMappingRequest(BaseModel):
    """Request body for admin edits to the column mapping schema."""

    standard_name: str
    synonym: str


# ---------------------------------------------------------------------------
# Scoped SugMosad (institution type) apply models
# ---------------------------------------------------------------------------

class SelectedRowsRequest(BaseModel):
    """One SugMosad value applied to a set of rows identified by _row_uid.

    Mirrors the row-selection mechanism used by row deletion: the UI sends
    the stable _row_uid values for the rows the user has selected in the grid.
    """
    sug_mosad: str
    row_uids: List[str]

    @field_validator("sug_mosad")
    @classmethod
    def validate_sug_mosad(cls, v: str) -> str:
        return _validate_numeric_min3(v, "סוג מוסד")

    @field_validator("row_uids")
    @classmethod
    def validate_row_uids(cls, v: List[str]) -> List[str]:
        if not v:
            raise ValueError("row_uids לא יכול להיות ריק")
        return v


class ApplySugMosadRequest(BaseModel):
    """Request body for POST /mosad-type/apply-scoped.

    scope:
        "workbook"      — apply sug_mosad to all sheets.
        "sheet"         — apply sug_mosad to sheet_name only.
        "selected_rows" — apply each entry in selected_rows (up to 3) to the
                          matching rows (by _row_uid) inside sheet_name.

    Validation:
        - sug_mosad must be numeric-only, ≥ 3 digits (for workbook/sheet scope).
        - mosad_id, if provided, is stored as-is after trimming.
        - sheet_name required for "sheet" and "selected_rows" scopes.
        - selected_rows required (1–3 entries) for "selected_rows" scope.
    """
    scope: str                                      # "workbook" | "sheet" | "selected_rows"
    sug_mosad: Optional[str] = None                 # used for workbook / sheet scope
    mosad_id: Optional[str] = None                  # optional workbook-level MosadID update
    sheet_name: Optional[str] = None                # required for sheet / selected_rows scope
    selected_rows: Optional[List[SelectedRowsRequest]] = None  # required for selected_rows scope

    @field_validator("scope")
    @classmethod
    def validate_scope(cls, v: str) -> str:
        allowed = {"workbook", "sheet", "selected_rows"}
        if v not in allowed:
            raise ValueError(f"scope חייב להיות אחד מ: {', '.join(sorted(allowed))}")
        return v

    @field_validator("sug_mosad")
    @classmethod
    def validate_sug_mosad(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return v
        return _validate_numeric_min3(v, "סוג מוסד")


    @model_validator(mode="after")
    def validate_scope_fields(self) -> "ApplySugMosadRequest":
        if self.scope in ("sheet", "selected_rows") and not self.sheet_name:
            raise ValueError("sheet_name נדרש עבור scope 'sheet' ו-'selected_rows'")
        if self.scope == "selected_rows":
            if not self.selected_rows:
                raise ValueError("selected_rows נדרש עבור scope 'selected_rows'")
            if len(self.selected_rows) > 3:
                raise ValueError("ניתן להגדיר עד 3 קבוצות שורות")
        if self.scope in ("workbook", "sheet") and not self.sug_mosad:
            raise ValueError(f"sug_mosad נדרש עבור scope '{self.scope}'")
        return self
