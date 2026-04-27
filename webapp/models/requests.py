"""Pydantic request models for the API layer."""

from typing import List
from pydantic import BaseModel


class CellEditRequest(BaseModel):
    """Request body for editing a single cell value."""
    row_uid: str
    field_name: str
    new_value: str


class DeleteRowRequest(BaseModel):
    """Request body for deleting one or more rows from a sheet."""
    row_uids: List[str]
