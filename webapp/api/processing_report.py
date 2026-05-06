"""Processing report router."""

from fastapi import APIRouter, Depends, Query

from webapp.dependencies import get_processing_report_service
from webapp.models.processing_report import ProcessingReport
from webapp.services.processing_report_service import ProcessingReportService

router = APIRouter(tags=["processing-report"])


@router.get(
    "/workbook/{session_id}/processing-report",
    response_model=ProcessingReport,
    response_model_exclude_none=True,
)
def get_processing_report(
    session_id: str,
    include_details: bool = Query(
        False,
        description="Include row-level validation diagnostics. Defaults to compact summaries only.",
    ),
    report_service: ProcessingReportService = Depends(get_processing_report_service),
) -> ProcessingReport:
    """Return the latest non-sensitive processing report for a session."""
    return report_service.get(session_id, include_details=include_details)
