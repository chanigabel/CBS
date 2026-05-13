"""Service layer that runs the standardization pipeline for a session."""

import logging
from datetime import date
from typing import List, Optional

from fastapi import HTTPException

from src.excel_standardization.processing.standardization_pipeline import StandardizationPipeline
from src.excel_standardization.engines.name_engine import NameEngine
from src.excel_standardization.engines.gender_engine import GenderEngine
from src.excel_standardization.engines.date_engine import DateEngine
from src.excel_standardization.engines.identifier_engine import IdentifierEngine
from src.excel_standardization.engines.text_processor import TextProcessor
from src.excel_standardization.data_types import SheetDataset

from webapp.models.responses import StandardizeResponse, PerSheetStat
from webapp.services.processing_report_service import ProcessingReportService
from webapp.services.session_service import SessionService
from webapp.services.workbook_loader import (
    extract_sheet_dataset,
    extract_workbook_dataset,
    get_workbook_sheet_names,
)

logger = logging.getLogger(__name__)


# שירות runtime שמריץ את StandardizationPipeline על session פעיל.
class StandardizationService:
    """Runs the standardization pipeline on a session's working copy."""

    # מקבל שירותי session/report כדי לעדכן מצב ודוחות במהלך העיבוד.
    def __init__(
        self,
        session_service: SessionService,
        processing_report_service: ProcessingReportService | None = None,
    ) -> None:
        self.session_service = session_service
        self.processing_report_service = (
            processing_report_service or ProcessingReportService(session_service)
        )

    # מחלץ מחדש גיליון או workbook, מריץ pipeline ומחזיר סטטיסטיקות ל־UI.
    def standardize(self, session_id: str, sheet_name: Optional[str] = None) -> StandardizeResponse:
        """Run standardization on the session's working copy.

        If *sheet_name* is given, only that sheet is (re-)standardized and the
        rest of the in-memory dataset is left untouched.  This is the fast path
        used by the UI.  When *sheet_name* is None all loaded sheets are
        processed (kept for CLI / batch compatibility).
        """
        record = self.session_service.get(session_id)

        pipeline = self._build_pipeline()

        loaded_workbook_now = False
        if record.workbook_dataset is None:
            # Load the workbook when no sheet has been accessed yet.
            try:
                wbd = extract_workbook_dataset(record.working_copy_path)
                self.session_service.update(session_id, workbook_dataset=wbd)
                self.processing_report_service.complete_stage(session_id, "extract")
                self.processing_report_service.update_workbook_counts(session_id, wbd)
                record = self.session_service.get(session_id)
                loaded_workbook_now = True
            except Exception as exc:
                self.processing_report_service.add_error(
                    session_id,
                    "Failed to extract workbook for standardization.",
                )
                logger.error(
                    "standardization_extract_failed",
                    exc_info=True,
                    extra={
                        "event": "standardization_extract_failed",
                        "session_id": session_id,
                        "error_type": type(exc).__name__,
                    },
                )
                raise HTTPException(
                    status_code=500,
                    detail="No workbook data available. Please load a sheet first.",
                )

        # Determine which sheets to re-extract and normalize.
        if sheet_name is not None:
            # Single-sheet path.
            try:
                existing = record.workbook_dataset.get_sheet_by_name(sheet_name)
                fresh = extract_sheet_dataset(record.working_copy_path, sheet_name)
                mosad_id = existing.get_metadata("MosadID") if existing is not None else None
                if mosad_id is not None:
                    fresh.set_metadata("MosadID", mosad_id)
                sheets_to_normalize = [fresh]
                self.processing_report_service.complete_stage(session_id, "extract")
            except KeyError:
                raise HTTPException(
                    status_code=404,
                    detail=f"Sheet '{sheet_name}' not found.",
                )
            except HTTPException:
                raise
            except Exception as exc:
                self.processing_report_service.add_error(
                    session_id,
                    f"Failed to extract sheet '{sheet_name}'.",
                )
                logger.error(
                    "standardization_sheet_extract_failed",
                    exc_info=True,
                    extra={
                        "event": "standardization_sheet_extract_failed",
                        "session_id": session_id,
                        "sheet_name": sheet_name,
                        "error_type": type(exc).__name__,
                    },
                )
                raise HTTPException(
                    status_code=500,
                    detail="Failed to read the working copy for standardization.",
                )
        else:
            # Full-workbook path.
            try:
                if loaded_workbook_now:
                    sheets_to_normalize = list(record.workbook_dataset.sheets)
                else:
                    sheets_to_normalize = []
                    for sname in get_workbook_sheet_names(record.working_copy_path):
                        fresh = extract_sheet_dataset(record.working_copy_path, sname)
                        existing = record.workbook_dataset.get_sheet_by_name(sname)
                        mosad_id = existing.get_metadata("MosadID") if existing is not None else None
                        if mosad_id is not None:
                            fresh.set_metadata("MosadID", mosad_id)
                        sheets_to_normalize.append(fresh)
                self.processing_report_service.complete_stage(session_id, "extract")
            except Exception as exc:
                self.processing_report_service.add_error(
                    session_id,
                    "Failed to extract workbook for standardization.",
                )
                logger.error(
                    "standardization_extract_failed",
                    exc_info=True,
                    extra={
                        "event": "standardization_extract_failed",
                        "session_id": session_id,
                        "error_type": type(exc).__name__,
                    },
                )
                raise HTTPException(
                    status_code=500,
                    detail="Failed to read the working copy for standardization.",
                )

        # Normalize
        normalized_sheets: List[SheetDataset] = []
        per_sheet_stats: List[PerSheetStat] = []
        failed_sheets: List[str] = []

        for sheet in sheets_to_normalize:
            try:
                norm = pipeline.normalize_dataset(sheet)
                normalized_sheets.append(norm)
                stats = norm.get_metadata("standardization_statistics", {})
                per_sheet_stats.append(PerSheetStat(
                    sheet_name=sheet.sheet_name,
                    rows=stats.get("total_rows", len(norm.rows)),
                    success_rate=stats.get("success_rate", 1.0),
                ))
                logger.info(
                    "sheet_standardized",
                    extra={
                        "event": "sheet_standardized",
                        "session_id": session_id,
                        "sheet_name": sheet.sheet_name,
                        "rows": per_sheet_stats[-1].rows,
                    },
                )
            except Exception as exc:
                self.processing_report_service.add_warning(
                    session_id,
                    f"Standardization failed for sheet '{sheet.sheet_name}'.",
                )
                logger.error(
                    "sheet_standardization_failed",
                    exc_info=True,
                    extra={
                        "event": "sheet_standardization_failed",
                        "session_id": session_id,
                        "sheet_name": sheet.sheet_name,
                        "error_type": type(exc).__name__,
                    },
                )
                failed_sheets.append(sheet.sheet_name)

        if not normalized_sheets:
            self.processing_report_service.add_error(
                session_id,
                "Standardization failed for all sheets.",
            )
            raise HTTPException(
                status_code=500,
                detail=f"standardization failed for all sheets: {', '.join(failed_sheets)}",
            )

        # Merge normalized sheets back into the session dataset.
        norm_by_name = {s.sheet_name: s for s in normalized_sheets}
        updated_sheets = []
        for existing in record.workbook_dataset.sheets:
            if existing.sheet_name in norm_by_name:
                updated_sheets.append(norm_by_name.pop(existing.sheet_name))
            else:
                updated_sheets.append(existing)
        # Any newly normalized sheets not previously in the dataset
        updated_sheets.extend(norm_by_name.values())
        record.workbook_dataset.sheets = updated_sheets

        # Workbook-level institution-report validation runs after normalization.
        try:
            from src.excel_standardization.validation.institution_report_validator import (
                InstitutionReportValidator,
                KNOWN_SHEETS,
            )
            from src.excel_standardization.services.sheet_name_resolver import (
                resolve_canonical_sheet_name,
            )

            sheets_for_validation = {
                resolve_canonical_sheet_name(s.sheet_name): s.rows
                for s in record.workbook_dataset.sheets
                if resolve_canonical_sheet_name(s.sheet_name) in KNOWN_SHEETS
            }
            # Pass sheet metadata so the validator can see MosadID and SugMosad.
            sheet_meta_for_validation = {
                resolve_canonical_sheet_name(s.sheet_name): {
                    "MosadID": record.mosad_id or s.get_metadata("MosadID"),
                    "SugMosad": record.mosad_types[0] if record.mosad_types else None,
                }
                for s in record.workbook_dataset.sheets
                if resolve_canonical_sheet_name(s.sheet_name) in KNOWN_SHEETS
            }
            if sheets_for_validation:
                wv = InstitutionReportValidator()
                wv.validate_workbook(
                    sheets_for_validation,
                    sheet_metadata=sheet_meta_for_validation,
                )
                logger.debug(
                    "Workbook-level institution-report validation completed for session %s",
                    session_id,
                )
            self.processing_report_service.complete_stage(session_id, "validate")
        except Exception as _wv_exc:
            self.processing_report_service.add_warning(
                session_id,
                "Workbook-level validation was skipped.",
            )
            logger.warning(
                "Workbook-level institution-report validation skipped: %s", _wv_exc
            )

        # Replay manual edits recorded before standardization.
        if record.edits:
            for (edit_sheet, edit_row_uid, edit_field), edit_value in record.edits.items():
                sheet_obj = record.workbook_dataset.get_sheet_by_name(edit_sheet)
                if sheet_obj is None:
                    continue
                for row in sheet_obj.rows:
                    if row.get("_row_uid") == edit_row_uid and edit_field in row:
                        row[edit_field] = edit_value
                        break
            logger.debug(
                f"Replayed {len(record.edits)} manual edit(s) after standardization "
                f"for session {session_id}"
            )

        self.session_service.update(session_id, status="standardized")

        total_rows = sum(s.rows for s in per_sheet_stats)
        record = self.session_service.get(session_id)
        self.processing_report_service.complete_stage(session_id, "standardize")
        self.processing_report_service.update_workbook_counts(session_id, record.workbook_dataset)
        logger.info(
            "standardization_complete",
            extra={
                "event": "standardization_complete",
                "session_id": session_id,
                "sheets_processed": len(normalized_sheets),
                "rows_processed": total_rows,
            },
        )

        return StandardizeResponse(
            session_id=session_id,
            status="standardized",
            sheets_processed=len(normalized_sheets),
            total_rows=total_rows,
            per_sheet_stats=per_sheet_stats,
        )

    # Backward-compatible alias
    # alias תפעולי ל־standardize עבור endpoints או קריאות קיימות.
    def normalize(self, session_id: str, sheet_name: Optional[str] = None) -> StandardizeResponse:
        return self.standardize(session_id, sheet_name=sheet_name)

    # בונה pipeline עם כל המנועים שה־Web flow מפעיל בפועל.
    def _build_pipeline(self) -> StandardizationPipeline:
        """Build a StandardizationPipeline with the active engines."""
        tp = TextProcessor()
        reference_date = date.today()
        return StandardizationPipeline(
            name_engine=NameEngine(tp),
            gender_engine=GenderEngine(),
            date_engine=DateEngine(reference_date=reference_date),
            identifier_engine=IdentifierEngine(),
            apply_name_standardization_enabled=True,
            apply_gender_standardization_enabled=True,
            apply_date_standardization_enabled=True,
            apply_identifier_standardization_enabled=True,
            reference_date=reference_date,
        )


# Backward-compatible alias for callers that still import the legacy name.
standardizationService = StandardizationService
