"""UploadService: handles file upload validation, storage, and session creation."""

import logging
import shutil
from pathlib import Path
from uuid import uuid4

from fastapi import HTTPException

from webapp.models.responses import UploadResponse
from webapp.models.session import SessionRecord
from webapp.services.processing_report_service import ProcessingReportService
from webapp.services.session_service import SessionService
from webapp.services.workbook_loader import (
    ALLOWED_WORKBOOK_EXTENSIONS,
    WorkbookLoadError,
    get_workbook_sheet_names,
)

logger = logging.getLogger(__name__)

ALLOWED_EXTENSIONS = ALLOWED_WORKBOOK_EXTENSIONS


# שירות העלאה שמייצר session ועותק עבודה לקובץ ה־Excel.
class UploadService:
    """Handles file upload: validation, saving to disk, and session creation."""

    # מקבל תיקיית העלאות ושירותי session/report הנדרשים לזרימה.
    def __init__(
        self,
        session_service: SessionService,
        uploads_dir: Path,
        work_dir: Path,
        processing_report_service: ProcessingReportService | None = None,
    ) -> None:
        self.session_service = session_service
        self.uploads_dir = uploads_dir
        self.work_dir = work_dir
        self.processing_report_service = (
            processing_report_service or ProcessingReportService(session_service)
        )

    # שומר את הקובץ, מחלץ שמות גיליונות ופותח רשומת session חדשה.
    def handle_upload(self, filename: str, file_bytes: bytes) -> UploadResponse:
        """Process an uploaded file and create a new session.

        Args:
            filename: Original filename from the upload
            file_bytes: Raw bytes of the uploaded file

        Returns:
            UploadResponse with session_id and sheet_names

        Raises:
            HTTPException 400: If file extension is not .xlsx, .xlsm, or .xls
            HTTPException 422: If file cannot be opened as a valid Excel workbook
            HTTPException 500: If an IO error occurs while saving the file
        """
        logger.info(
            "upload_started",
            extra={
                "event": "upload_started",
                "upload_filename": filename,
                "upload_size_bytes": len(file_bytes),
            },
        )
        # 1. Validate extension
        suffix = Path(filename).suffix.lower()
        if suffix not in ALLOWED_EXTENSIONS:
            logger.warning(
                "upload_rejected_invalid_extension",
                extra={
                    "event": "upload_rejected_invalid_extension",
                    "upload_filename": filename,
                    "extension": suffix,
                },
            )
            raise HTTPException(
                status_code=400,
                detail=(
                    f"File format not supported. "
                    f"Please upload a .xlsx, .xlsm, or .xls file. Got: '{suffix}'"
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
            logger.info(
                "upload_saved_internal_files",
                extra={
                    "event": "upload_saved_internal_files",
                    "session_id": session_id,
                    "extension": suffix,
                    "upload_size_bytes": len(file_bytes),
                },
            )
        except Exception as exc:
            logger.exception(
                "upload_save_failed",
                extra={
                    "event": "upload_save_failed",
                    "session_id": session_id,
                    "error_type": type(exc).__name__,
                },
            )
            raise HTTPException(
                status_code=500,
                detail="Failed to save the uploaded file. Please try again.",
            )

        # 5. Validate workbook and get sheet names through the shared loader.
        # Full per-sheet extraction is deferred to the first sheet load request.
        try:
            sheet_names = get_workbook_sheet_names(working_path)
        except WorkbookLoadError as exc:
            source_path.unlink(missing_ok=True)
            working_path.unlink(missing_ok=True)
            logger.warning(
                "upload_rejected_invalid_workbook",
                extra={
                    "event": "upload_rejected_invalid_workbook",
                    "session_id": session_id,
                    "upload_filename": filename,
                    "error_type": type(exc).__name__,
                },
            )
            raise HTTPException(status_code=422, detail=str(exc))
        except Exception as exc:
            source_path.unlink(missing_ok=True)
            working_path.unlink(missing_ok=True)
            logger.warning(
                "upload_rejected_invalid_workbook",
                extra={
                    "event": "upload_rejected_invalid_workbook",
                    "session_id": session_id,
                    "upload_filename": filename,
                    "error_type": type(exc).__name__,
                },
            )
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
        logger.info(
            "upload_session_created",
            extra={
                "event": "upload_session_created",
                "session_id": session_id,
                "sheet_count": len(sheet_names),
            },
        )
        self.processing_report_service.start(session_id)
        self.processing_report_service.complete_stage(session_id, "upload")

        sheet_names = list(sheet_names)
        logger.info(
            "upload_successful",
            extra={
                "event": "upload_successful",
                "session_id": session_id,
                "upload_filename": filename,
                "sheet_count": len(sheet_names),
            },
        )

        return UploadResponse(session_id=session_id, sheet_names=sheet_names)
