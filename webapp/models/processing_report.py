"""Processing report model shared by UI and API-only flows."""

from typing import Any, List, Optional

from pydantic import BaseModel, Field


class MissingInputColumnsBySheet(BaseModel):
    sheet_name: str
    columns: List[str] = Field(default_factory=list)


class MissingRequiredExportField(BaseModel):
    sheet_name: str
    field_name: str
    rows_affected: int


class SummaryCount(BaseModel):
    message: str
    count: int


class ReportSummary(BaseModel):
    """Summary-level statistics for the entire processing run."""
    total_rows: int = 0
    rows_processed: int = 0
    rows_exported: int = 0
    rows_changed_automatically: int = 0
    rows_changed_manually: int = 0
    sheets_with_warnings: int = 0
    sheets_completed: int = 0
    sheets_completed_with_warnings: int = 0
    sheets_failed: int = 0


class MissingRequiredFieldSummary(BaseModel):
    field: str
    count: int


class InvalidDateValue(BaseModel):
    sheet_name: str
    row_number: Optional[int] = None
    row_uid: Optional[str] = None
    source_field: str
    raw_value: Any = None
    corrected_value: Any = None
    status_message: str


class InvalidIdentifierValue(BaseModel):
    sheet_name: str
    row_number: Optional[int] = None
    row_uid: Optional[str] = None
    source_field: str
    raw_value: Any = None
    corrected_value: Any = None
    status_message: str


class PerSheetProcessingReport(BaseModel):
    sheet_name: str
    total_rows: int = 0
    rows_processed: int = 0
    rows_exported: int = 0
    rows_changed_automatically: int = 0
    rows_changed_manually: int = 0
    sheet_status: str = "unknown"  # "בוצע" | "בוצע עם אזהרות" | "נכשל" | "unknown"
    warning_counts: dict = Field(default_factory=dict)  # e.g., {"invalid_date": 5, "invalid_id": 2}
    warnings: List[str] = Field(default_factory=list)
    errors: List[str] = Field(default_factory=list)


class ProcessingReport(BaseModel):
    """Non-sensitive processing summary for a workbook session."""

    session_id: str
    status: str = "success"
    status_reason: str = "success"
    completed_stages: List[str] = Field(default_factory=list)
    sheets_processed: int = 0
    rows_processed: int = 0
    rows_exported: int = 0
    rows_changed_automatically: int = 0
    rows_changed_manually: int = 0
    missing_input_columns: List[MissingInputColumnsBySheet] = Field(default_factory=list)
    identifier_summary: List[SummaryCount] = Field(default_factory=list)
    date_summary: List[SummaryCount] = Field(default_factory=list)
    missing_required_fields: List[MissingRequiredFieldSummary] = Field(default_factory=list)
    empty_required_columns_summary: List[MissingRequiredFieldSummary] = Field(default_factory=list)
    missing_required_export_fields: List[MissingRequiredExportField] = Field(default_factory=list)
    invalid_date_values: Optional[List[InvalidDateValue]] = None
    invalid_identifier_values: Optional[List[InvalidIdentifierValue]] = None
    per_sheet_warnings: List[PerSheetProcessingReport] = Field(default_factory=list)
    summary: Optional[ReportSummary] = None
    warnings: List[str] = Field(default_factory=list)
    errors: List[str] = Field(default_factory=list)
    output_filename: str = ""
    duration: float = 0.0
