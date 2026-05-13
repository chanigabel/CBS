"""Workbook service for session-backed workbook summaries and sheet data."""

import logging
import uuid
from fastapi import HTTPException

from webapp.models.responses import SheetDataResponse, SheetSummary, WorkbookSummary
from webapp.services.grid_payload import build_sheet_grid_payload
from webapp.services.session_service import SessionService
from webapp.services.workbook_loader import (
    WorkbookLoadError,
    extract_sheet_dataset,
    get_workbook_sheet_names,
    sheet_exists,
    workbook_suffix,
)

logger = logging.getLogger(__name__)


class WorkbookService:
    """Provides workbook summary and sheet data from in-memory session state."""

    def __init__(self, session_service: SessionService) -> None:
        self.session_service = session_service

    def _ensure_sheet_loaded(self, record, sheet_name: str) -> None:
        """Lazily extract a single sheet from disk if not yet in the dataset."""
        from src.excel_standardization.data_types import WorkbookDataset

        working_path = record.working_copy_path
        is_xls = workbook_suffix(working_path) == ".xls"

        if record.workbook_dataset is not None:
            if record.workbook_dataset.get_sheet_by_name(sheet_name) is not None:
                return
            if not sheet_exists(working_path, sheet_name):
                raise HTTPException(
                    status_code=404,
                    detail=f"Sheet '{sheet_name}' not found in this workbook.",
                )

        try:
            sheet_dataset = extract_sheet_dataset(working_path, sheet_name)
        except KeyError:
            raise HTTPException(
                status_code=404,
                detail=f"Sheet '{sheet_name}' not found in this workbook.",
            )
        except WorkbookLoadError as exc:
            logger.error("Failed to extract sheet '%s': %s", sheet_name, exc, exc_info=True)
            raise HTTPException(status_code=422 if not is_xls else 500, detail=str(exc))

        if record.workbook_dataset is None:
            try:
                all_names = get_workbook_sheet_names(working_path)
            except Exception:
                all_names = [sheet_name]
            record.workbook_dataset = WorkbookDataset(
                source_file=working_path,
                sheets=[sheet_dataset],
                metadata={"sheet_names": list(all_names)},
            )
        else:
            record.workbook_dataset.sheets.append(sheet_dataset)

    def get_summary(self, session_id: str) -> WorkbookSummary:
        """Return a summary of all sheets in the workbook."""
        record = self.session_service.get(session_id)

        if record.workbook_dataset is None:
            try:
                names = get_workbook_sheet_names(record.working_copy_path)
            except WorkbookLoadError as exc:
                raise HTTPException(status_code=422, detail=str(exc))
            except Exception:
                raise HTTPException(
                    status_code=500,
                    detail="Workbook data is not available for this session.",
                )
            sheets = [
                SheetSummary(sheet_name=n, row_count=0, field_names=[])
                for n in names
            ]
            return WorkbookSummary(session_id=session_id, sheets=sheets)

        sheets = [
            SheetSummary(
                sheet_name=sheet.sheet_name,
                row_count=sheet.get_row_count(),
                field_names=sheet.get_field_names(),
            )
            for sheet in record.workbook_dataset.sheets
        ]
        return WorkbookSummary(session_id=session_id, sheets=sheets)

    def get_sheet_data(self, session_id: str, sheet_name: str) -> SheetDataResponse:
        """Return all rows for a specific sheet."""
        record = self.session_service.get(session_id)
        self._ensure_sheet_loaded(record, sheet_name)

        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)
        if sheet is None:
            raise HTTPException(
                status_code=404,
                detail=f"Sheet '{sheet_name}' not found in this workbook.",
            )

        for row in sheet.rows:
            if "_row_uid" not in row:
                row["_row_uid"] = uuid.uuid4().hex

        session_mosad_id = record.mosad_id or None
        meta_mosad_id = session_mosad_id or sheet.get_metadata("MosadID")
        active_mosad_type = record.mosad_types[0] if record.mosad_types else None
        return build_sheet_grid_payload(
            sheet,
            session_mosad_id=session_mosad_id or "",
            active_mosad_type=active_mosad_type,
            metadata_mosad_id=meta_mosad_id,
        )
