"""Upload router: POST /api/upload"""

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile

from webapp.dependencies import get_upload_service
from webapp.models.responses import UploadResponse
from webapp.services.upload_service import UploadService

router = APIRouter(tags=["upload"])

# F-05: Maximum accepted upload size (50 MB).  Checked after reading the file
# bytes so we can return a clear 413 rather than letting a huge file silently
# exhaust server memory.
_MAX_UPLOAD_BYTES = 50 * 1024 * 1024  # 50 MB


@router.post("/upload", response_model=UploadResponse)
async def upload_file(
    file: UploadFile = File(...),
    upload_service: UploadService = Depends(get_upload_service),
) -> UploadResponse:
    """Upload an Excel workbook and create a new session.

    Accepts .xlsx or .xlsm files up to 50 MB.
    Returns a session_id and list of sheet names.
    """
    file_bytes = await file.read()

    # F-05: Reject files that exceed the size limit before any further processing.
    if len(file_bytes) > _MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=413,
            detail=(
                f"File too large ({len(file_bytes) // (1024 * 1024)} MB). "
                f"Maximum allowed size is {_MAX_UPLOAD_BYTES // (1024 * 1024)} MB."
            ),
        )

    return upload_service.handle_upload(file.filename or "upload.xlsx", file_bytes)
