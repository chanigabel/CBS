"""User-facing workbook processing report models."""

from __future__ import annotations

from typing import Dict, List

from pydantic import BaseModel, Field


class ReportSummary(BaseModel):
    total_sheets: int = 0
    total_rows: int = 0
    edited_cells: int = 0
    rows_with_warnings: int = 0
    rows_with_errors: int = 0
    rows_without_issues: int = 0
    corrected_fields: int = 0


class ManualEditsSummary(BaseModel):
    edited_cells: int = 0
    edited_sheets: List[str] = Field(default_factory=list)
    edited_fields: List[str] = Field(default_factory=list)


class ReportIssue(BaseModel):
    sheet_name: str
    row_uid: str = ""
    row_number: int = 0
    field_name: str = ""
    status_field: str = ""
    status_message: str = ""
    severity: str = "warning"


class SheetReport(BaseModel):
    sheet_name: str
    row_count: int = 0
    column_count: int = 0
    rows_with_warnings: int = 0
    rows_with_errors: int = 0
    corrected_fields: int = 0
    issues_count: int = 0
    status_counts: Dict[str, Dict[str, int]] = Field(default_factory=dict)


class WorkbookProcessingReport(BaseModel):
    session_id: str
    file_name: str = ""
    status: str = ""
    export_ready: bool = False
    dirty: bool = False
    stale: bool = False
    export_blocked_reason: str = ""
    summary: ReportSummary = Field(default_factory=ReportSummary)
    manual_edits: ManualEditsSummary = Field(default_factory=ManualEditsSummary)
    sheets: List[SheetReport] = Field(default_factory=list)
    issues: List[ReportIssue] = Field(default_factory=list)
