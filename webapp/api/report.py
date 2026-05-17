"""User-facing workbook report router."""

from fastapi import APIRouter, Depends, Query
from fastapi.responses import FileResponse

from webapp.api.export import _content_disposition
from webapp.dependencies import get_report_export_service, get_report_service
from webapp.models.report import WorkbookProcessingReport
from webapp.services.report_export_service import ReportExportService
from webapp.services.report_service import ReportService

router = APIRouter(tags=["report"])


@router.get(
    "/workbook/{session_id}/report",
    response_model=WorkbookProcessingReport,
    response_model_exclude_none=True,
)
def get_workbook_report(
    session_id: str,
    include_details: bool = Query(
        False,
        description="Include row-level issues in the report response.",
    ),
    report_service: ReportService = Depends(get_report_service),
) -> WorkbookProcessingReport:
    """Return a read-only user-facing report for the current workbook state."""
    return report_service.build(session_id, include_details=include_details)


@router.get(
    "/workbook/{session_id}/report/export",
)
def export_workbook_report(
    session_id: str,
    report_export_service: ReportExportService = Depends(get_report_export_service),
) -> FileResponse:
    """Export the processing report as a separate downloadable workbook."""
    output_path = report_export_service.export(session_id)
    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_path.name,
        headers={"Content-Disposition": _content_disposition(output_path.name)},
    )
