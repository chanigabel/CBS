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
from webapp.services.column_mapping_schema import ColumnMappingSchemaService
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

    def __init__(
        self,
        session_service: SessionService,
        column_schema_service: ColumnMappingSchemaService | None = None,
    ) -> None:
        self.session_service = session_service
        self.column_schema_service = column_schema_service

    def validate_column_mappings_before_standardization(
        self,
        session_id: str,
        sheet_name: str | None = None,
    ) -> None:
        """Block standardization when duplicate effective standard targets exist."""
        record = self.session_service.get(session_id)

        if record.workbook_dataset is None:
            return

        duplicate_messages: list[str] = []

        for sheet in record.workbook_dataset.sheets:
            if sheet_name is not None and sheet.sheet_name != sheet_name:
                continue

            mappings = record.column_mappings.get(sheet.sheet_name, {})

            for target_name, source_names in self._target_to_sources(sheet, mappings).items():
                if len(source_names) > 1:
                    duplicate_messages.append(
                        f"Cannot start standardization. In sheet '{sheet.sheet_name}', "
                        f"the field '{target_name}' is mapped to more than one column "
                        f"({', '.join(source_names)}). Please fix the column mapping so "
                        "each standard field appears only once."
                    )

        if duplicate_messages:
            raise HTTPException(status_code=400, detail=" ".join(duplicate_messages))

    def prepare_column_mappings_for_standardization(
        self,
        session_id: str,
        sheet_name: str | None = None,
    ) -> None:
        """Validate and apply stored column mappings immediately before standardization."""
        record = self.session_service.get(session_id)

        if record.workbook_dataset is None:
            return

        self.validate_column_mappings_before_standardization(
            session_id,
            sheet_name=sheet_name,
        )

        for sheet in record.workbook_dataset.sheets:
            if sheet_name is not None and sheet.sheet_name != sheet_name:
                continue

            mappings = record.column_mappings.get(sheet.sheet_name, {})
            if mappings:
                self.apply_column_mappings_to_sheet(sheet, mappings)
                record.column_mappings.pop(sheet.sheet_name, None)

        if not record.column_mappings:
            record.column_mappings = {}

    def _ensure_sheet_loaded(self, record, sheet_name: str) -> None:
        """Lazily extract a single sheet from disk if not yet in the dataset."""
        from src.excel_standardization.data_types import WorkbookDataset

        working_path = record.working_copy_path
        is_xls = workbook_suffix(working_path) == ".xls"
        session_id = getattr(record, "session_id", "")

        if record.workbook_dataset is not None:
            if record.workbook_dataset.get_sheet_by_name(sheet_name) is not None:
                return
            if not sheet_exists(working_path, sheet_name):
                logger.warning(
                    "sheet_load_rejected_not_found",
                    extra={
                        "event": "sheet_load_rejected_not_found",
                        "session_id": session_id,
                        "sheet_name": sheet_name,
                    },
                )
                raise HTTPException(
                    status_code=404,
                    detail=f"Sheet '{sheet_name}' not found in this workbook.",
                )

        try:
            logger.info(
                "sheet_load_started",
                extra={
                    "event": "sheet_load_started",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                    "lazy_extraction": True,
                },
            )
            sheet_dataset = extract_sheet_dataset(working_path, sheet_name)
        except KeyError:
            logger.warning(
                "sheet_load_rejected_not_found",
                extra={
                    "event": "sheet_load_rejected_not_found",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                },
            )
            raise HTTPException(
                status_code=404,
                detail=f"Sheet '{sheet_name}' not found in this workbook.",
            )
        except WorkbookLoadError as exc:
            logger.exception(
                "sheet_load_failed",
                extra={
                    "event": "sheet_load_failed",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                    "error_type": type(exc).__name__,
                },
            )
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

        logger.info(
            "sheet_load_completed",
            extra={
                "event": "sheet_load_completed",
                "session_id": session_id,
                "sheet_name": sheet_name,
                "row_count": len(sheet_dataset.rows),
                "column_count": len(sheet_dataset.field_names),
                "lazy_extraction": True,
            },
        )

    def get_column_schema(self) -> ColumnSchemaResponse:
        """Return the supported generic target field names for column mapping."""
        if self.column_schema_service is None:
            return ColumnSchemaResponse(fields=[])

        return ColumnSchemaResponse(
            fields=self.column_schema_service.fields(),
            mappings=self.column_schema_service.mappings(),
            suggestions=self.column_schema_service.suggestions(),
        )

    def _resolve_explicit_mapping_target(self, target_name: str) -> str:
        """Resolve only a target chosen explicitly by the user."""
        target_name = (target_name or "").strip()

        if self.column_schema_service is not None:
            return self.column_schema_service.resolve(target_name)

        return target_name

    def _is_supported_standard_field(self, field_name: str) -> bool:
        """Return True only for already-standard field names.

        Important:
        This must not call resolve(), because source/display columns such as
        'מס_סידורי' are not standardized fields and should not fail validation.
        """
        if not field_name:
            return False

        if self.column_schema_service is None:
            return True

        return field_name in set(self.column_schema_service.fields())

    def _target_to_sources(self, sheet, mappings: dict) -> dict[str, list[str]]:
        """Build target-to-source mapping for duplicate validation.

        Only explicit user mappings are resolved against the schema.
        Unmapped non-standard columns are ignored for duplicate-standard-field validation.
        """
        target_to_sources: dict[str, list[str]] = {}

        for source_name in sheet.field_names:
            if not source_name or source_name.startswith("_"):
                continue

            if source_name in mappings:
                target_name = self._resolve_explicit_mapping_target(mappings[source_name])
            elif self._is_supported_standard_field(source_name):
                target_name = source_name
            else:
                continue

            target_to_sources.setdefault(target_name, []).append(source_name)

        return target_to_sources

    def apply_column_mappings_to_sheet(self, sheet, mappings: dict) -> None:
        """Apply source-to-standard field mappings safely and atomically.

        Unmapped columns are preserved unchanged.
        Only explicit mapping targets are resolved against the standardized schema.
        """
        if not mappings:
            return

        original_field_names = list(sheet.field_names)
        target_by_source: dict[str, str] = {}

        for source_name in original_field_names:
            if source_name in mappings:
                target_by_source[source_name] = self._resolve_explicit_mapping_target(
                    mappings[source_name]
                )
            else:
                target_by_source[source_name] = source_name

        new_field_names = [target_by_source[field] for field in original_field_names]

        duplicates = sorted(
            target for target in set(new_field_names) if new_field_names.count(target) > 1
        )
        if duplicates:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"Cannot apply column mappings in sheet '{sheet.sheet_name}'. "
                    f"Duplicate target field(s): {', '.join(duplicates)}."
                ),
            )

        new_rows = []

        for row in sheet.rows:
            new_row = {}

            for source_name in original_field_names:
                target_name = target_by_source[source_name]

                if source_name in row:
                    if target_name in new_row:
                        raise HTTPException(
                            status_code=400,
                            detail=(
                                f"Cannot apply column mappings in sheet '{sheet.sheet_name}'. "
                                f"Field '{target_name}' would be overwritten."
                            ),
                        )

                    new_row[target_name] = row[source_name]

            for key, value in row.items():
                if key not in target_by_source:
                    if key in new_row:
                        raise HTTPException(
                            status_code=400,
                            detail=(
                                f"Cannot apply column mappings in sheet '{sheet.sheet_name}'. "
                                f"Field '{key}' would be overwritten."
                            ),
                        )

                    new_row[key] = value

            new_rows.append(new_row)

        sheet.field_names = new_field_names

        for row, new_row in zip(sheet.rows, new_rows):
            row.clear()
            row.update(new_row)

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

        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)  # type: ignore[union-attr]
        if sheet is None:
            logger.warning(
                "sheet_data_not_found_after_load",
                extra={
                    "event": "sheet_data_not_found_after_load",
                    "session_id": session_id,
                    "sheet_name": sheet_name,
                },
            )
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

        logger.info(
            "sheet_data_returned",
            extra={
                "event": "sheet_data_returned",
                "session_id": session_id,
                "sheet_name": sheet_name,
                "row_count": len(sheet.rows),
                "column_count": len(sheet.field_names),
            },
        )

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
        """Persist a source-to-standard column mapping without mutating row data."""
        old_name = (old_name or "").strip()
        new_name = (new_name or "").strip()

        if not old_name or not new_name:
            raise HTTPException(
                status_code=400,
                detail="old_name and new_name are required.",
            )

        new_name = self._resolve_explicit_mapping_target(new_name)

        record = self.session_service.get(session_id)
        self._ensure_sheet_loaded(record, sheet_name)

        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)  # type: ignore[union-attr]
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

        mappings = record.column_mappings.setdefault(sheet_name, {})
        before = dict(mappings)

        if old_name == new_name:
            mappings.pop(old_name, None)
        else:
            mappings[old_name] = new_name

        if not mappings:
            record.column_mappings.pop(sheet_name, None)

        current = record.column_mappings.get(sheet_name, {})
        if record.status == "standardized" and before != current:
            record.working_dataset_dirty = True

        return ColumnMappingResponse(
            sheet_name=sheet_name,
            old_name=old_name,
            new_name=new_name,
            field_names=list(sheet.field_names),
            column_mappings=dict(current),
        )

    def reload_column_mapping(
        self,
        session_id: str,
        sheet_name: str,
    ) -> ColumnSchemaResponse:
        """Reload central mapping config without mutating the loaded sheet data."""
        if self.column_schema_service is not None:
            self.column_schema_service.reload()

        record = self.session_service.get(session_id)
        self._ensure_sheet_loaded(record, sheet_name)

        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)  # type: ignore[union-attr]
        if sheet is None:
            raise HTTPException(
                status_code=404,
                detail=f"Sheet '{sheet_name}' not found in this workbook.",
            )

        return self.get_column_schema()