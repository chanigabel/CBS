"""Shared ProcessingReportService for UI and API-only flows."""

from __future__ import annotations

import logging
import time
from typing import Iterable

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
from webapp.services.report_aggregation import (
    aggregate_identifier_messages as _aggregate_identifier_messages,
    aggregate_missing_required_fields as _aggregate_missing_required_fields,
    aggregate_validation_messages as _aggregate_validation_messages,
)
from webapp.services.report_collectors import (
    aggregate_detail_messages_by_sheet,
    aggregate_identifier_messages_by_sheet,
    build_per_sheet_warnings,
    collect_invalid_date_values,
    collect_invalid_identifier_values,
    collect_missing_input_columns,
    collect_missing_required_export_fields,
    empty_required_columns_message_for_sheet,
)
from webapp.services.report_status_builder import (
    is_invalid_date_component,
    is_real_identifier_issue,
    refresh_status,
    row_number,
    status_reason,
)
from webapp.services.session_service import SessionService

logger = logging.getLogger(__name__)


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
        if stage not in ["upload", "extract", "standardize", "validate", "export"]:
            raise ValueError(f"Unknown processing stage: {stage}")
        if stage not in report.completed_stages:
            report.completed_stages.append(stage)
            report.completed_stages.sort(key=["upload", "extract", "standardize", "validate", "export"].index)
        refresh_status(report)
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
        report.missing_input_columns = collect_missing_input_columns(workbook_dataset)
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
        refresh_status(report)
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
        report.missing_input_columns = collect_missing_input_columns(record.workbook_dataset)
        report.missing_required_export_fields = collect_missing_required_export_fields(record)
        report.empty_required_columns_summary = _aggregate_missing_required_fields(
            report.missing_required_export_fields
        )
        report.missing_required_fields = list(report.empty_required_columns_summary)
        invalid_date_values = collect_invalid_date_values(record)
        invalid_identifier_values = collect_invalid_identifier_values(record, is_real_identifier_issue)
        report.invalid_date_values = invalid_date_values
        report.invalid_identifier_values = invalid_identifier_values
        report.date_summary = _aggregate_validation_messages(
            item.status_message for item in invalid_date_values
        )
        report.identifier_summary = _aggregate_identifier_messages(
            invalid_identifier_values,
            is_real_identifier_issue,
        )
        report.per_sheet_warnings = build_per_sheet_warnings(
            record,
            rows_exported_by_sheet,
            report,
        )
        refresh_status(report)
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
        refresh_status(report)
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
        refresh_status(report)
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
        refresh_status(report)
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
        return _aggregate_missing_required_fields(collect_missing_required_export_fields(record))

    def aggregate_missing_required_fields(
        self,
        fields: Iterable[MissingRequiredExportField],
    ) -> list[MissingRequiredFieldSummary]:
        return _aggregate_missing_required_fields(fields)

    def aggregate_validation_messages(self, messages: Iterable[str]) -> list[SummaryCount]:
        return _aggregate_validation_messages(messages)

    def aggregate_identifier_messages(
        self,
        details: Iterable[InvalidIdentifierValue],
    ) -> list[SummaryCount]:
        return _aggregate_identifier_messages(details, is_real_identifier_issue)

    def collect_missing_required_export_fields(self, record) -> list[MissingRequiredExportField]:
        return collect_missing_required_export_fields(record)

    def collect_missing_input_columns(self, workbook_dataset) -> list[MissingInputColumnsBySheet]:
        return collect_missing_input_columns(workbook_dataset)

    def collect_invalid_date_values(self, record) -> list[InvalidDateValue]:
        return collect_invalid_date_values(record)

    def collect_invalid_identifier_values(self, record) -> list[InvalidIdentifierValue]:
        return collect_invalid_identifier_values(record, is_real_identifier_issue)

    def build_per_sheet_warnings(
        self,
        record,
        rows_exported_by_sheet: dict[str, int],
        report: ProcessingReport,
    ) -> list[PerSheetProcessingReport]:
        return build_per_sheet_warnings(record, rows_exported_by_sheet, report)

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


# Backward-compatible aliases for direct imports in older tests/code
aggregate_missing_required_fields = _aggregate_missing_required_fields
aggregate_validation_messages = _aggregate_validation_messages
aggregate_identifier_messages = _aggregate_identifier_messages

