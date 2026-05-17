"""Export service facade for writing the normalized workbook to disk."""

from __future__ import annotations

import logging
import os
from pathlib import Path

from fastapi import HTTPException
from openpyxl import Workbook

from webapp.services.export_rows import build_export_filename as _build_export_filename
from webapp.services.export_rows import resolve_sug_mosad_for_sheet as _resolve_sug_mosad_for_sheet
from webapp.services.export_rows import visible_rows
from webapp.services.export_schema import EXPORT_MAPPING, canonical_sheet_name, headers_for_sheet
from webapp.services.export_writer import write_export_workbook
from webapp.services.processing_report_service import ProcessingReportService
from webapp.services.session_service import SessionService

logger = logging.getLogger(__name__)


# שירות יצוא שמייצר קובץ Excel להורדה מתוך ה־WorkbookDataset שב־session.
class ExportService:
    """Writes the current in-memory WorkbookDataset to an Excel file for download."""

    # מקבל session/output/report כדי לכתוב קובץ ולעדכן דוח עיבוד.
    def __init__(
        self,
        session_service: SessionService,
        output_dir: Path,
        processing_report_service: ProcessingReportService | None = None,
    ) -> None:
        self.session_service = session_service
        self.output_dir = output_dir
        self.processing_report_service = (
            processing_report_service or ProcessingReportService(session_service)
        )

    # מוודא שיש Dataset, כותב קובץ יצוא ומחזיר את נתיב ההורדה.
    def export(self, session_id: str) -> Path:
        """Export the session's workbook using the fixed export schema."""
        logger.info(
            "export_requested",
            extra={"event": "export_requested", "session_id": session_id},
        )
        record = self.session_service.get(session_id)

        # If the session was already standardized, allow manual edits (e.g.
        # user-corrected *_corrected fields) without blocking export. Keep a
        # warning in the log when the working dataset is dirty, but do not
        # prevent export after a successful standardization run.
        if record.working_dataset_dirty:
            logger.info(
                "export_dirty_dataset_allowed",
                extra={
                    "event": "export_dirty_dataset_allowed",
                    "session_id": session_id,
                    "working_dataset_dirty": True,
                },
            )

        workbook_dataset = record.workbook_dataset
        if (
            record.status != "standardized"
            or workbook_dataset is None
            or not getattr(workbook_dataset, "sheets", None)
        ):
            raise HTTPException(
                status_code=409,
                detail="Run Standardization before exporting. Export uses the latest successful Standardization result.",
            )

        output_filename = _build_export_filename(record)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        logger.info(
            "export_started",
            extra={
                "event": "export_started",
                "session_id": session_id,
                "output_filename": output_filename,
                "sheet_count": len(record.workbook_dataset.sheets),
                "row_count": sum(len(sheet.rows) for sheet in record.workbook_dataset.sheets),
            },
        )

        original_stem = Path(record.original_filename).stem
        for old_file in self.output_dir.glob(f"{original_stem}_standardized_*.xlsx"):
            try:
                old_file.unlink()
                logger.debug("Removed previous export: %s", old_file.name)
            except Exception as exc:
                logger.warning("Could not remove old export file %s: %s", old_file, exc)

        output_path = self.output_dir / output_filename
        temp_path = output_path.with_name(f"{output_path.name}.tmp")
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                logger.warning("Could not remove stale temp export file %s", temp_path, exc_info=True)

        try:
            rows_exported, rows_exported_by_sheet = write_export_workbook(
                record,
                temp_path,
                workbook_factory=Workbook,
            )
            if output_path.exists():
                try:
                    output_path.unlink()
                except Exception:
                    logger.warning("Could not remove previous export file %s", output_path, exc_info=True)
            os.replace(temp_path, output_path)
            self.processing_report_service.finalize_export_details(
                session_id,
                record=record,
                rows_exported_by_sheet=rows_exported_by_sheet,
                output_filename=output_path.name,
            )
            logger.info(
                "export_successful",
                extra={
                    "event": "export_successful",
                    "session_id": session_id,
                    "output_filename": output_path.name,
                    "rows_exported": rows_exported,
                },
            )
        except Exception as exc:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    logger.warning("Could not remove failed temp export file %s", temp_path, exc_info=True)
            self.processing_report_service.add_error(session_id, "Export failed.")
            logger.error(
                "export_failed",
                exc_info=True,
                extra={
                    "event": "export_failed",
                    "session_id": session_id,
                    "error_type": type(exc).__name__,
                },
            )
            raise HTTPException(
                status_code=500,
                detail="Export failed. Please try again. Your session data is preserved.",
            )

        return output_path


# Backward-compatible aliases for callers and tests that still import helpers
canonical_sheet_name = canonical_sheet_name
headers_for_sheet = headers_for_sheet
EXPORT_MAPPING = EXPORT_MAPPING
_build_export_filename = _build_export_filename
_resolve_sug_mosad_for_sheet = _resolve_sug_mosad_for_sheet
visible_rows = visible_rows
