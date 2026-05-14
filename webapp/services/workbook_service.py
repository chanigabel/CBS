"""Workbook service for session-backed workbook summaries and sheet data."""

import logging
import uuid
from fastapi import HTTPException

from webapp.models.responses import (
    ColumnMappingResponse,
    ColumnSchemaResponse,
    SheetDataResponse,
    SheetSummary,
    WorkbookSummary,
)
from webapp.services.grid_payload import build_sheet_grid_payload
from webapp.services.column_mapping_schema import ColumnMappingSchemaService
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

    def __init__(
        self,
        session_service: SessionService,
        column_schema_service: ColumnMappingSchemaService | None = None,
    ) -> None:
        self.session_service = session_service
        self.column_schema_service = column_schema_service

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

        mappings = record.column_mappings.get(sheet_name, {})
        if mappings:
            self.apply_column_mappings_to_sheet(sheet_dataset, mappings)

    def get_column_schema(self) -> ColumnSchemaResponse:
        """Return the supported generic target field names for column mapping."""
        if self.column_schema_service is None:
            return ColumnSchemaResponse(fields=[])
        return ColumnSchemaResponse(
            fields=self.column_schema_service.fields(),
            mappings=self.column_schema_service.mappings(),
            suggestions=self.column_schema_service.suggestions(),
        )

    def apply_column_mappings_to_sheet(self, sheet, mappings: dict) -> None:
        """Apply stored source-to-standard field mappings to a SheetDataset."""
        for old_name, new_name in mappings.items():
            if old_name == new_name:
                continue
            if old_name not in sheet.field_names:
                continue
            if new_name in sheet.field_names and new_name != old_name:
                continue
            sheet.field_names = [
                new_name if field == old_name else field
                for field in sheet.field_names
            ]
            for row in sheet.rows:
                if old_name in row:
                    row[new_name] = row.pop(old_name)

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
            column_mappings=record.column_mappings.get(sheet_name, {}),
        )

    def update_column_mapping(
        self,
        session_id: str,
        sheet_name: str,
        old_name: str,
        new_name: str,
    ) -> ColumnMappingResponse:
        """Rename a loaded source column and persist the mapping for normalization."""
        old_name = (old_name or "").strip()
        new_name = (new_name or "").strip()
        if not old_name or not new_name:
            raise HTTPException(status_code=400, detail="old_name and new_name are required.")
        if self.column_schema_service is not None:
            new_name = self.column_schema_service.resolve(new_name)

        record = self.session_service.get(session_id)
        self._ensure_sheet_loaded(record, sheet_name)
        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)
        if sheet is None:
            raise HTTPException(
                status_code=404,
                detail=f"Sheet '{sheet_name}' not found in this workbook.",
            )
        if old_name not in sheet.field_names:
            raise HTTPException(
                status_code=404,
                detail=f"Column '{old_name}' not found in sheet '{sheet_name}'.",
            )
        if old_name == new_name:
            mappings = record.column_mappings.setdefault(sheet_name, {})
            return ColumnMappingResponse(
                sheet_name=sheet_name,
                old_name=old_name,
                new_name=new_name,
                field_names=list(sheet.field_names),
                column_mappings=dict(mappings),
            )
        if new_name in sheet.field_names:
            raise HTTPException(
                status_code=409,
                detail=f"Column '{new_name}' already exists in sheet '{sheet_name}'.",
            )

        sheet.field_names = [
            new_name if field == old_name else field
            for field in sheet.field_names
        ]
        for row in sheet.rows:
            if old_name in row:
                row[new_name] = row.pop(old_name)

        mappings = record.column_mappings.setdefault(sheet_name, {})
        mappings[old_name] = new_name

        # Manual cell edits are keyed by field name; keep those references valid.
        updated_edits = {}
        for (edit_sheet, row_uid, field_name), value in record.edits.items():
            if edit_sheet == sheet_name and field_name == old_name:
                updated_edits[(edit_sheet, row_uid, new_name)] = value
            else:
                updated_edits[(edit_sheet, row_uid, field_name)] = value
        record.edits = updated_edits

        return ColumnMappingResponse(
            sheet_name=sheet_name,
            old_name=old_name,
            new_name=new_name,
            field_names=list(sheet.field_names),
            column_mappings=dict(mappings),
        )

    def reload_column_mapping(
        self,
        session_id: str,
        sheet_name: str,
    ) -> ColumnSchemaResponse:
        """Reload central mapping config and re-apply stored mappings to a sheet."""
        if self.column_schema_service is not None:
            self.column_schema_service.reload()
        record = self.session_service.get(session_id)
        self._ensure_sheet_loaded(record, sheet_name)
        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)
        if sheet is None:
            raise HTTPException(
                status_code=404,
                detail=f"Sheet '{sheet_name}' not found in this workbook.",
            )
        mappings = record.column_mappings.get(sheet_name, {})
        if mappings:
            self.apply_column_mappings_to_sheet(sheet, mappings)
        return self.get_column_schema()
