"""Process-file router: upload, standardize, and export in one request."""

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from webapp.api.export import _content_disposition
from webapp.api.upload import _MAX_UPLOAD_BYTES
from webapp.dependencies import (
    get_export_service,
    get_standardization_service,
    get_upload_service,
)
from webapp.services.export_service import ExportService
from webapp.services.standardization_service import standardizationService
from webapp.services.upload_service import UploadService

router = APIRouter(tags=["process-file"])


@router.post("/process-file")
async def process_file(
    file: UploadFile = File(...),
    upload_service: UploadService = Depends(get_upload_service),
    standardization_service: standardizationService = Depends(get_standardization_service),
    export_service: ExportService = Depends(get_export_service),
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

    upload_response = upload_service.handle_upload(file.filename or "upload.xlsx", file_bytes)
    standardization_service.standardize(upload_response.session_id)
    output_path = export_service.export(upload_response.session_id)

    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_path.name,
        headers={"Content-Disposition": _content_disposition(output_path.name)},
    )
