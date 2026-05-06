"""Shared ProcessingReportService for UI and API-only flows."""

from __future__ import annotations

import logging
import time
from collections import Counter, defaultdict
from typing import Iterable, Optional

from fastapi import HTTPException

from webapp.models.processing_report import (
    InvalidDateValue,
    InvalidIdentifierValue,
    MissingInputColumnsBySheet,
    MissingRequiredFieldSummary,
    MissingRequiredExportField,
    PerSheetProcessingReport,
    ProcessingReport,
    SummaryCount,
)
from webapp.services.session_service import SessionService

logger = logging.getLogger(__name__)

_STAGE_ORDER = ["upload", "extract", "standardize", "validate", "export"]

_EXPECTED_BASE_INPUT_COLUMNS = [
    "first_name",
    "last_name",
    "father_name",
    "id_number",
    "passport",
    "gender",
]

_DATE_COMPONENTS = [
    "birth_year",
    "birth_month",
    "birth_day",
    "entry_year",
    "entry_month",
    "entry_day",
]


class ProcessingReportService:
    """Maintains non-sensitive processing reports on session records."""

    def __init__(self, session_service: SessionService) -> None:
        self.session_service = session_service
        self._started_at: dict[str, float] = {}

    def start(self, session_id: str) -> ProcessingReport:
        report = ProcessingReport(session_id=session_id)
        self._started_at[session_id] = time.perf_counter()
        self._save(session_id, report)
        logger.info(
            "processing_report_started",
            extra={"event": "processing_report_started", "session_id": session_id},
        )
        return report

    def get(self, session_id: str, include_details: bool = False) -> ProcessingReport:
        record = self.session_service.get(session_id)
        if record.processing_report is None:
            raise HTTPException(status_code=404, detail="Processing report not found.")
        if include_details:
            return record.processing_report
        report = record.processing_report.model_copy(deep=True)
        report.invalid_date_values = None
        report.invalid_identifier_values = None
        return report

    def complete_stage(self, session_id: str, stage: str) -> ProcessingReport:
        report = self._ensure(session_id)
        if stage not in _STAGE_ORDER:
            raise ValueError(f"Unknown processing stage: {stage}")
        if stage not in report.completed_stages:
            report.completed_stages.append(stage)
            report.completed_stages.sort(key=_STAGE_ORDER.index)
        self._refresh_status(report)
        self._touch_duration(session_id, report)
        self._save(session_id, report)
        logger.info(
            "processing_stage_completed",
            extra={
                "event": "processing_stage_completed",
                "session_id": session_id,
                "stage": stage,
                "status": report.status,
            },
        )
        return report

    def update_workbook_counts(self, session_id: str, workbook_dataset) -> ProcessingReport:
        report = self._ensure(session_id)
        report.sheets_processed = len(workbook_dataset.sheets)
        report.rows_processed = sum(len(sheet.rows) for sheet in workbook_dataset.sheets)
        report.missing_input_columns = self.collect_missing_input_columns(workbook_dataset)
        report.per_sheet_warnings = self._merge_per_sheet_reports(
            report.per_sheet_warnings,
            {
                sheet.sheet_name: {
                    "rows_processed": len(sheet.rows),
                    "rows_exported": None,
                    "warnings": [],
                    "errors": [],
                }
                for sheet in workbook_dataset.sheets
            },
        )
        self._refresh_status(report)
        self._touch_duration(session_id, report)
        self._save(session_id, report)
        logger.info(
            "processing_counts_updated",
            extra={
                "event": "processing_counts_updated",
                "session_id": session_id,
                "sheets_processed": report.sheets_processed,
                "rows_processed": report.rows_processed,
            },
        )
        return report

    def mark_exported(
        self,
        session_id: str,
        rows_exported: int,
        output_filename: str,
    ) -> ProcessingReport:
        report = self._ensure(session_id)
        report.rows_exported = rows_exported
        report.output_filename = output_filename
        return self.complete_stage(session_id, "export")

    def finalize_export_details(
        self,
        session_id: str,
        record,
        rows_exported_by_sheet: dict[str, int],
        output_filename: str,
    ) -> ProcessingReport:
        report = self._ensure(session_id)
        report.output_filename = output_filename
        report.rows_exported = sum(rows_exported_by_sheet.values())
        report.missing_input_columns = self.collect_missing_input_columns(record.workbook_dataset)
        report.missing_required_export_fields = self.collect_missing_required_export_fields(record)
        report.empty_required_columns_summary = self.aggregate_missing_required_fields(
            report.missing_required_export_fields
        )
        report.missing_required_fields = list(report.empty_required_columns_summary)
        invalid_date_values = self.collect_invalid_date_values(record)
        invalid_identifier_values = self.collect_invalid_identifier_values(record)
        report.invalid_date_values = invalid_date_values
        report.invalid_identifier_values = invalid_identifier_values
        report.date_summary = self.aggregate_validation_messages(
            item.status_message for item in invalid_date_values
        )
        report.identifier_summary = self.aggregate_identifier_messages(invalid_identifier_values)
        report.per_sheet_warnings = self.build_per_sheet_warnings(
            record,
            rows_exported_by_sheet,
            report,
        )
        self._refresh_status(report)
        self._touch_duration(session_id, report)
        self._save(session_id, report)
        self.complete_stage(session_id, "export")
        logger.info(
            "processing_report_details_finalized",
            extra={
                "event": "processing_report_details_finalized",
                "session_id": session_id,
                "missing_input_sheet_count": len(report.missing_input_columns),
                "missing_required_field_count": len(report.missing_required_export_fields),
                "invalid_date_value_count": len(invalid_date_values),
                "invalid_identifier_value_count": len(invalid_identifier_values),
            },
        )
        return self.get(session_id)

    def set_missing_required_fields(
        self,
        session_id: str,
        missing_required_fields: Iterable[str],
    ) -> ProcessingReport:
        report = self._ensure(session_id)
        report.missing_required_fields = [
            MissingRequiredFieldSummary(field=field, count=1)
            for field in dict.fromkeys(missing_required_fields)
        ]
        self._refresh_status(report)
        self._touch_duration(session_id, report)
        self._save(session_id, report)
        if report.missing_required_fields:
            logger.warning(
                "processing_missing_required_fields",
                extra={
                    "event": "processing_missing_required_fields",
                    "session_id": session_id,
                    "missing_required_field_count": len(report.missing_required_fields),
                },
            )
        return report

    def add_warning(self, session_id: str, message: str) -> ProcessingReport:
        report = self._ensure(session_id)
        if message not in report.warnings:
            report.warnings.append(message)
        self._refresh_status(report)
        self._touch_duration(session_id, report)
        self._save(session_id, report)
        logger.warning(
            "processing_warning",
            extra={
                "event": "processing_warning",
                "session_id": session_id,
                "report_message": message,
            },
        )
        return report

    def add_error(self, session_id: str, message: str) -> ProcessingReport:
        report = self._ensure(session_id)
        if message not in report.errors:
            report.errors.append(message)
        self._refresh_status(report)
        self._touch_duration(session_id, report)
        self._save(session_id, report)
        logger.error(
            "processing_error",
            extra={
                "event": "processing_error",
                "session_id": session_id,
                "report_message": message,
            },
        )
        return report

    def collect_missing_required_fields(self, record) -> list[MissingRequiredFieldSummary]:
        """Return aggregate missing required field counts."""
        return self.aggregate_missing_required_fields(
            self.collect_missing_required_export_fields(record)
        )

    def aggregate_missing_required_fields(
        self,
        fields: Iterable[MissingRequiredExportField],
    ) -> list[MissingRequiredFieldSummary]:
        counts = Counter()
        for field in fields:
            counts[field.field_name] += field.rows_affected
        return [
            MissingRequiredFieldSummary(field=field_name, count=count)
            for field_name, count in sorted(counts.items())
        ]

    def aggregate_validation_messages(self, messages: Iterable[str]) -> list[SummaryCount]:
        counts = Counter(message for message in messages if message)
        return [
            SummaryCount(message=message, count=count)
            for message, count in sorted(counts.items())
        ]

    def aggregate_identifier_messages(
        self,
        details: Iterable[InvalidIdentifierValue],
    ) -> list[SummaryCount]:
        counts = Counter()
        seen_rows = set()
        for item in details:
            if not self._is_real_identifier_issue(item.status_message):
                continue
            row_key = item.row_uid if item.row_uid is not None else item.row_number
            key = (item.sheet_name, row_key, item.status_message)
            if key in seen_rows:
                continue
            seen_rows.add(key)
            counts[item.status_message] += 1
        return [
            SummaryCount(message=message, count=count)
            for message, count in sorted(counts.items())
        ]

    def collect_missing_required_export_fields(self, record) -> list[MissingRequiredExportField]:
        """Return aggregate missing export-field details without row values."""
        if record.workbook_dataset is None:
            return []

        from webapp.services.export_service import (
            EXPORT_MAPPING,
            _resolve_sug_mosad_for_sheet,
            canonical_sheet_name,
            headers_for_sheet,
            visible_rows,
        )

        missing = Counter()
        active_mosad_type = record.mosad_types[0] if record.mosad_types else ""

        for sheet_dataset in record.workbook_dataset.sheets:
            export_name = canonical_sheet_name(sheet_dataset.sheet_name)
            schema = headers_for_sheet(export_name)
            data_rows, _ui_cols = visible_rows(sheet_dataset)
            scoped_type = _resolve_sug_mosad_for_sheet(
                record.sug_mosad_configs,
                sheet_dataset.sheet_name,
                active_mosad_type,
            )

            for row in data_rows:
                effective_row = dict(row)
                if record.mosad_id:
                    effective_row["MosadID"] = record.mosad_id
                if callable(scoped_type):
                    scoped_value = scoped_type(row.get("_row_uid", ""))
                    if scoped_value is not None:
                        effective_row["SugMosad"] = scoped_value
                elif scoped_type:
                    effective_row["SugMosad"] = scoped_type

                for header in schema:
                    json_key = EXPORT_MAPPING.get(header)
                    if json_key is None:
                        continue
                    value = effective_row.get(json_key)
                    if value is None or str(value).strip() == "":
                        missing[(export_name, header)] += 1

        return [
            MissingRequiredExportField(
                sheet_name=sheet,
                field_name=field,
                rows_affected=count,
            )
            for (sheet, field), count in sorted(missing.items())
        ]

    def collect_missing_input_columns(self, workbook_dataset) -> list[MissingInputColumnsBySheet]:
        if workbook_dataset is None:
            return []

        result = []
        for sheet in workbook_dataset.sheets:
            fields = set(sheet.field_names or [])
            expected = list(_EXPECTED_BASE_INPUT_COLUMNS)

            if "birth_date" not in fields:
                expected.extend(["birth_year", "birth_month", "birth_day"])
            if "entry_date" not in fields:
                expected.extend(["entry_year", "entry_month", "entry_day"])

            missing = [field for field in expected if field not in fields]
            if missing:
                result.append(MissingInputColumnsBySheet(sheet_name=sheet.sheet_name, columns=missing))

        return result

    def collect_invalid_date_values(self, record) -> list[InvalidDateValue]:
        if record.workbook_dataset is None:
            return []

        invalid = []
        for sheet in record.workbook_dataset.sheets:
            for idx, row in enumerate(sheet.rows):
                row_number = self._row_number(sheet, idx)
                row_uid = row.get("_row_uid")
                for field in _DATE_COMPONENTS:
                    if field not in row:
                        continue
                    prefix = "birth" if field.startswith("birth_") else "entry"
                    status = row.get(f"{prefix}_date_status") or ""
                    corrected = row.get(f"{field}_corrected")
                    if not self._is_invalid_date_component(status, corrected):
                        continue
                    invalid.append(
                        InvalidDateValue(
                            sheet_name=sheet.sheet_name,
                            row_number=row_number if row_uid is None else None,
                            row_uid=row_uid,
                            source_field=field,
                            raw_value=row.get(field),
                            corrected_value=corrected,
                            status_message=status,
                        )
                    )

        return invalid

    def collect_invalid_identifier_values(self, record) -> list[InvalidIdentifierValue]:
        if record.workbook_dataset is None:
            return []

        invalid = []
        for sheet in record.workbook_dataset.sheets:
            for idx, row in enumerate(sheet.rows):
                status = row.get("identifier_status") or ""
                if not self._is_real_identifier_issue(status):
                    continue
                row_number = self._row_number(sheet, idx)
                row_uid = row.get("_row_uid")
                for field in ("id_number", "passport"):
                    if field not in row:
                        continue
                    invalid.append(
                        InvalidIdentifierValue(
                            sheet_name=sheet.sheet_name,
                            row_number=row_number if row_uid is None else None,
                            row_uid=row_uid,
                            source_field=field,
                            raw_value=row.get(field),
                            corrected_value=row.get(f"{field}_corrected"),
                            status_message=status,
                        )
                    )

        return invalid

    def build_per_sheet_warnings(
        self,
        record,
        rows_exported_by_sheet: dict[str, int],
        report: ProcessingReport,
    ) -> list[PerSheetProcessingReport]:
        if record.workbook_dataset is None:
            return []

        missing_input_by_sheet = {item.sheet_name: len(item.columns) for item in report.missing_input_columns}
        invalid_dates_by_sheet = self._aggregate_detail_messages_by_sheet(report.invalid_date_values or [])
        invalid_ids_by_sheet = self._aggregate_identifier_messages_by_sheet(
            report.invalid_identifier_values or []
        )

        per_sheet = []
        for sheet in record.workbook_dataset.sheets:
            warnings = []
            if missing_input_by_sheet.get(sheet.sheet_name):
                warnings.append(f"{missing_input_by_sheet[sheet.sheet_name]} expected input column(s) missing")
            export_name = self._export_name_for_sheet(sheet.sheet_name)
            required_columns_message = self._empty_required_columns_message_for_sheet(
                export_name,
                report.missing_required_export_fields,
            )
            if required_columns_message:
                warnings.append(required_columns_message)
            for message, count in invalid_dates_by_sheet.get(sheet.sheet_name, []):
                warnings.append(f"{message}: {count} date value(s)")
            for message, count in invalid_ids_by_sheet.get(sheet.sheet_name, []):
                warnings.append(f"{message}: {count} identifier value(s)")

            per_sheet.append(
                PerSheetProcessingReport(
                    sheet_name=sheet.sheet_name,
                    rows_processed=len(sheet.rows),
                    rows_exported=rows_exported_by_sheet.get(sheet.sheet_name, 0),
                    warnings=warnings,
                    errors=[],
                )
            )

        return per_sheet

    def _ensure(self, session_id: str) -> ProcessingReport:
        record = self.session_service.get(session_id)
        if record.processing_report is None:
            return self.start(session_id)
        return record.processing_report

    def _save(self, session_id: str, report: ProcessingReport) -> None:
        self.session_service.update(session_id, processing_report=report)

    def _touch_duration(self, session_id: str, report: ProcessingReport) -> None:
        started_at = self._started_at.get(session_id)
        if started_at is not None:
            report.duration = round(time.perf_counter() - started_at, 3)

    def _refresh_status(self, report: ProcessingReport) -> None:
        if report.errors:
            report.status = "failed"
        elif (
            report.warnings
            or report.missing_required_fields
            or report.missing_input_columns
            or report.missing_required_export_fields
            or report.date_summary
            or report.identifier_summary
            or any(item.warnings for item in report.per_sheet_warnings)
        ):
            report.status = "partial_success"
        else:
            report.status = "success"
        report.status_reason = self._status_reason(report)

    def _status_reason(self, report: ProcessingReport) -> str:
        if report.status == "failed":
            return f"failed because {len(report.errors)} error(s) occurred"

        reasons = []
        empty_required_count = sum(item.count for item in report.empty_required_columns_summary)
        missing_identifier_count = sum(
            item.count for item in report.identifier_summary if "חסר מזהים" in item.message
        )
        invalid_identifier_count = sum(
            item.count for item in report.identifier_summary if "לא תקינה" in item.message
        )
        invalid_date_count = sum(item.count for item in report.date_summary)
        missing_input_count = sum(len(item.columns) for item in report.missing_input_columns)
        warning_count = len(report.warnings) + sum(len(item.warnings) for item in report.per_sheet_warnings)

        if missing_input_count:
            reasons.append(f"{missing_input_count} expected input columns are missing")
        if empty_required_count:
            reasons.append(f"{empty_required_count} ערכי חובה ריקים בקובץ הייצוא")
        if missing_identifier_count:
            reasons.append(f"{missing_identifier_count} rows missing identifiers")
        if invalid_date_count:
            reasons.append(f"{invalid_date_count} invalid date values")
        if invalid_identifier_count:
            reasons.append(f"{invalid_identifier_count} invalid ID values")
        if warning_count:
            reasons.append(f"{warning_count} warning(s) were reported")

        if report.status == "partial_success":
            return "partial_success because:\n- " + "\n- ".join(reasons or ["warnings were reported"])
        return "success because all processing stages completed without reportable issues"

    def _row_number(self, sheet, idx: int) -> int:
        return int(sheet.header_row or 1) + int(sheet.header_rows_count or 1) + idx + 1

    def _is_invalid_date_component(self, status: str, corrected) -> bool:
        if not status:
            return False
        invalid_status = (
            status == "ערך תאריך לא תקין"
            or "לא תקין" in status
            or "לא תקינה" in status
            or "לא קיים" in status
        )
        return invalid_status and (corrected is None or str(corrected).strip() == "")

    def _is_real_identifier_issue(self, status: str) -> bool:
        return bool(status) and ("חסר מזהים" in status or "לא תקינה" in status)

    def _export_name_for_sheet(self, sheet_name: str) -> str:
        from webapp.services.export_service import canonical_sheet_name

        return canonical_sheet_name(sheet_name)

    def _aggregate_detail_messages_by_sheet(self, details: Iterable) -> dict[str, list[tuple[str, int]]]:
        counts = defaultdict(Counter)
        for item in details:
            if item.status_message:
                counts[item.sheet_name][item.status_message] += 1
        return {
            sheet_name: sorted(message_counts.items())
            for sheet_name, message_counts in counts.items()
        }

    def _empty_required_columns_message_for_sheet(
        self,
        export_name: str,
        details: Iterable[MissingRequiredExportField],
    ) -> str:
        counts_by_field: dict[str, int] = {}
        for item in details:
            if item.sheet_name != export_name or item.rows_affected <= 0:
                continue
            counts_by_field[item.field_name] = counts_by_field.get(item.field_name, 0) + item.rows_affected

        if not counts_by_field:
            return ""

        from webapp.services.export_service import headers_for_sheet

        ordered_fields = []
        for field_name in headers_for_sheet(export_name):
            if field_name in counts_by_field:
                ordered_fields.append(f"{field_name}={counts_by_field[field_name]}")

        return f"עמודות חובה ריקות: {', '.join(ordered_fields)}"

    def _aggregate_identifier_messages_by_sheet(
        self,
        details: Iterable[InvalidIdentifierValue],
    ) -> dict[str, list[tuple[str, int]]]:
        counts = defaultdict(Counter)
        seen_rows = set()
        for item in details:
            if not item.status_message:
                continue
            row_key = item.row_uid if item.row_uid is not None else item.row_number
            key = (item.sheet_name, row_key, item.status_message)
            if key in seen_rows:
                continue
            seen_rows.add(key)
            counts[item.sheet_name][item.status_message] += 1
        return {
            sheet_name: sorted(message_counts.items())
            for sheet_name, message_counts in counts.items()
        }

    def _merge_per_sheet_reports(self, existing: list[PerSheetProcessingReport], updates: dict):
        by_name = {item.sheet_name: item for item in existing}
        for sheet_name, values in updates.items():
            current = by_name.get(sheet_name) or PerSheetProcessingReport(sheet_name=sheet_name)
            current.rows_processed = values.get("rows_processed", current.rows_processed)
            rows_exported = values.get("rows_exported")
            if rows_exported is not None:
                current.rows_exported = rows_exported
            current.warnings = values.get("warnings", current.warnings)
            current.errors = values.get("errors", current.errors)
            by_name[sheet_name] = current
        return list(by_name.values())
