"""UploadService: handles file upload validation, storage, and session creation."""

import logging
import shutil
from pathlib import Path
from uuid import uuid4

from fastapi import HTTPException

from webapp.models.responses import UploadResponse
from webapp.models.session import SessionRecord
from webapp.services.session_service import SessionService

logger = logging.getLogger(__name__)

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}


class UploadService:
    """Handles file upload: validation, saving to disk, and session creation."""

    def __init__(
        self,
        session_service: SessionService,
        uploads_dir: Path,
        work_dir: Path,
    ) -> None:
        self.session_service = session_service
        self.uploads_dir = uploads_dir
        self.work_dir = work_dir

    def handle_upload(self, filename: str, file_bytes: bytes) -> UploadResponse:
        """Process an uploaded file and create a new session.

        Args:
            filename: Original filename from the upload
            file_bytes: Raw bytes of the uploaded file

        Returns:
            UploadResponse with session_id and sheet_names

        Raises:
            HTTPException 400: If file extension is not .xlsx or .xlsm
            HTTPException 422: If file cannot be opened as a valid Excel workbook
            HTTPException 500: If an IO error occurs while saving the file
        """
        # 1. Validate extension
        suffix = Path(filename).suffix.lower()
        if suffix not in ALLOWED_EXTENSIONS:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"File format not supported. "
                    f"Please upload a .xlsx or .xlsm file. Got: '{suffix}'"
                ),
            )

        # 2. Generate session_id
        session_id = str(uuid4())

        # 3. Ensure directories exist
        self.uploads_dir.mkdir(parents=True, exist_ok=True)
        self.work_dir.mkdir(parents=True, exist_ok=True)

        # 4. Save source file (never modified)
        source_path = self.uploads_dir / f"{session_id}{suffix}"
        working_path = self.work_dir / f"{session_id}{suffix}"

        try:
            source_path.write_bytes(file_bytes)
            shutil.copy2(source_path, working_path)
        except Exception as exc:
            logger.error(f"Failed to save uploaded file: {exc}", exc_info=True)
            raise HTTPException(
                status_code=500,
                detail="Failed to save the uploaded file. Please try again.",
            )

        # 5. Validate workbook and get sheet names — open with openpyxl directly
        # to avoid a full extraction on upload.  Full per-sheet extraction is
        # deferred to the first sheet load request, keeping upload fast.
        try:
            from openpyxl import load_workbook as _load_wb
            _wb = _load_wb(str(working_path), data_only=True, read_only=True)
            sheet_names = _wb.sheetnames
            _wb.close()
            if not sheet_names:
                raise ValueError("Workbook has no sheets")
        except Exception as exc:
            source_path.unlink(missing_ok=True)
            working_path.unlink(missing_ok=True)
            logger.warning(f"Invalid workbook uploaded '{filename}': {exc}")
            raise HTTPException(
                status_code=422,
                detail=(
                    "The uploaded file could not be opened as a valid Excel workbook. "
                    "Please check the file and try again."
                ),
            )

        # 6. Create session with no workbook_dataset yet — it will be populated
        # lazily on the first GET /sheet request via WorkbookService.
        record = SessionRecord(
            session_id=session_id,
            source_file_path=str(source_path),
            working_copy_path=str(working_path),
            original_filename=filename,
            status="uploaded",
            workbook_dataset=None,
        )
        self.session_service.create(record)

        sheet_names = list(sheet_names)
        logger.info(
            f"Upload successful: session={session_id}, "
            f"file='{filename}', sheets={sheet_names}"
        )

        return UploadResponse(session_id=session_id, sheet_names=sheet_names)
