"""Process-file router: upload, standardize, and export in one request."""

import logging
from pathlib import Path

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from openpyxl import load_workbook

from webapp.api.export import _content_disposition
from webapp.api.upload import _MAX_UPLOAD_BYTES
from webapp.dependencies import (
    get_export_service,
    get_processing_report_service,
    get_standardization_service,
    get_upload_service,
)
from webapp.services.export_service import ExportService
from webapp.services.processing_report_service import ProcessingReportService
from webapp.services.standardization_service import standardizationService
from webapp.services.upload_service import UploadService

router = APIRouter(tags=["process-file"])
logger = logging.getLogger(__name__)


def _assert_real_excel_output(output_path: Path) -> None:
    """Verify that a generated response file is a readable Excel workbook."""
    if not output_path.exists() or output_path.stat().st_size == 0:
        raise HTTPException(
            status_code=500,
            detail="Processing did not produce an Excel output file.",
        )

    try:
        wb = load_workbook(str(output_path), read_only=True, data_only=True)
        wb.close()
    except Exception:
        output_path.unlink(missing_ok=True)
        raise HTTPException(
            status_code=500,
            detail="Processing produced an invalid Excel output file.",
        )


@router.post("/process-file")
async def process_file(
    file: UploadFile = File(...),
    upload_service: UploadService = Depends(get_upload_service),
    standardization_service: standardizationService = Depends(get_standardization_service),
    export_service: ExportService = Depends(get_export_service),
    report_service: ProcessingReportService = Depends(get_processing_report_service),
) -> FileResponse:
    """Upload, standardize, export, and return a workbook in one request."""
    file_bytes = await file.read()

    if len(file_bytes) > _MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=413,
            detail=(
                f"File too large ({len(file_bytes) // (1024 * 1024)} MB). "
                f"Maximum allowed size is {_MAX_UPLOAD_BYTES // (1024 * 1024)} MB."
            ),
        )

    session_id = None
    try:
        upload_response = upload_service.handle_upload(file.filename or "upload.xlsx", file_bytes)
        session_id = upload_response.session_id
        standardization_service.standardize(session_id)
        output_path = export_service.export(session_id)
        _assert_real_excel_output(output_path)
        report = report_service.get(session_id)
    except HTTPException as exc:
        if session_id is not None and exc.status_code >= 500:
            try:
                report_service.add_error(session_id, str(exc.detail))
            except Exception:
                pass
        raise
    except Exception as exc:
        if session_id is not None:
            try:
                report_service.add_error(session_id, "Processing failed.")
            except Exception:
                pass
        logger.error(
            "process_file_failed",
            exc_info=True,
            extra={
                "event": "process_file_failed",
                "session_id": session_id or "",
                "error_type": type(exc).__name__,
            },
        )
        raise HTTPException(
            status_code=500,
            detail="File processing failed. No Excel output was generated.",
        )

    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_path.name,
        headers={
            "Content-Disposition": _content_disposition(output_path.name),
            "X-Processing-Report-Id": upload_response.session_id,
            "X-Processing-Status": report.status,
        },
    )
