"""Build user-facing workbook processing reports from session state."""

from __future__ import annotations

from collections import Counter
from typing import Any, Iterable

from webapp.models.report import (
    ManualEditsSummary,
    ReportSummary,
    ReportIssue,
    SheetReport,
    WorkbookProcessingReport,
)
from webapp.services.session_service import SessionService

_EXPLICIT_STATUS_FIELDS = {
    "gender_status",
    "identifier_status",
    "birth_date_status",
    "entry_date_status",
    "_validation_status",
    "_standardization_failures",
}

_OK_STATUS_VALUES = {
    "",
    "ok",
    "valid",
    "success",
    "passed",
    "true",
    "תקין",
    "תקינה",
}

_ERROR_MARKERS = (
    "error",
    "failed",
    "failure",
    "invalid",
    "לא תקין",
    "לא תקינה",
    "שגיאה",
)


class ReportService:
    """Create a read-only business report from the current WorkbookDataset."""

    def __init__(self, session_service: SessionService) -> None:
        self.session_service = session_service

    def build(self, session_id: str, include_details: bool = False) -> WorkbookProcessingReport:
        record = self.session_service.get(session_id)
        report = WorkbookProcessingReport(
            session_id=session_id,
            file_name=record.original_filename or "",
            status=record.status,
            dirty=bool(record.working_dataset_dirty),
            stale=bool(record.working_dataset_dirty),
        )

        # Export is allowed only after at least one successful standardization
        # run. Manual edits after standardization should not block export; they
        # are reflected as `dirty` and `stale` in the report but do not prevent
        # users from exporting the latest standardized result with manual
        # corrections applied.
        if record.status == "standardized" and record.workbook_dataset is not None:
            report.export_ready = True
        else:
            report.export_ready = False
            if record.workbook_dataset is None:
                report.export_blocked_reason = "Workbook data is not loaded yet."
            else:
                report.export_blocked_reason = "Standardization has not completed yet."

        report.manual_edits = self._manual_edits_summary(record.edits)

        workbook_dataset = record.workbook_dataset
        if workbook_dataset is None:
            report.summary = ReportSummary(edited_cells=report.manual_edits.edited_cells)
            return report

        issues: list[ReportIssue] = []
        total_rows = 0
        total_warning_rows = 0
        total_error_rows = 0
        total_corrected_fields = 0
        sheets: list[SheetReport] = []

        for sheet in workbook_dataset.sheets:
            sheet_report, sheet_issues = self._sheet_report(sheet, include_details=include_details)
            sheets.append(sheet_report)
            if include_details:
                issues.extend(sheet_issues)
            total_rows += sheet_report.row_count
            total_warning_rows += sheet_report.rows_with_warnings
            total_error_rows += sheet_report.rows_with_errors
            total_corrected_fields += sheet_report.corrected_fields

        report.sheets = sheets
        report.summary = ReportSummary(
            total_sheets=len(sheets),
            total_rows=total_rows,
            edited_cells=report.manual_edits.edited_cells,
            rows_with_warnings=total_warning_rows,
            rows_with_errors=total_error_rows,
            rows_without_issues=max(total_rows - total_warning_rows - total_error_rows, 0),
            corrected_fields=total_corrected_fields,
        )
        if include_details:
            report.issues = issues
        return report

    def _sheet_report(self, sheet, include_details: bool = False) -> tuple[SheetReport, list[ReportIssue]]:
        status_counts: dict[str, Counter[str]] = {}
        warning_rows = 0
        error_rows = 0
        corrected_fields = 0
        issue_count = 0
        issues: list[ReportIssue] = []

        for row_number, row in enumerate(sheet.rows, start=1):
            has_warning = False
            has_error = False

            for field_name, raw_status in row.items():
                if not self._is_status_field(field_name):
                    continue
                values = self._status_values(raw_status)
                for value in values:
                    if self._is_ok_status(value):
                        continue
                    status_counts.setdefault(field_name, Counter())[value] += 1
                    severity = self._severity_for_status(field_name, value)
                    if severity == "error":
                        has_error = True
                    else:
                        has_warning = True
                    issue_count += 1
                    if include_details:
                        issues.append(
                            ReportIssue(
                                sheet_name=sheet.sheet_name,
                                row_uid=str(row.get("_row_uid") or ""),
                                row_number=row_number,
                                field_name=field_name,
                                status_field=field_name,
                                status_message=value,
                                severity=severity,
                            )
                        )

            corrected_fields += self._corrected_field_count(row)
            if has_error:
                error_rows += 1
            elif has_warning:
                warning_rows += 1

        return SheetReport(
            sheet_name=sheet.sheet_name,
            row_count=len(sheet.rows),
            column_count=len(sheet.field_names or []),
            rows_with_warnings=warning_rows,
            rows_with_errors=error_rows,
            corrected_fields=corrected_fields,
            issues_count=issue_count,
            status_counts={
                field: dict(sorted(counter.items()))
                for field, counter in sorted(status_counts.items())
            },
        ), issues

    @staticmethod
    def _manual_edits_summary(edits: dict) -> ManualEditsSummary:
        sheets = set()
        fields = set()
        for key in edits:
            if not isinstance(key, tuple) or len(key) != 3:
                continue
            sheet_name, _row_uid, field_name = key
            sheets.add(str(sheet_name))
            fields.add(str(field_name))
        return ManualEditsSummary(
            edited_cells=len(edits),
            edited_sheets=sorted(sheets),
            edited_fields=sorted(fields),
        )

    @staticmethod
    def _is_status_field(field_name: str) -> bool:
        return field_name in _EXPLICIT_STATUS_FIELDS or field_name.endswith("_status")

    @staticmethod
    def _status_values(raw_status: Any) -> list[str]:
        if raw_status is None:
            return []
        if isinstance(raw_status, (list, tuple, set)):
            return [str(value).strip() for value in raw_status if str(value).strip()]
        text = str(raw_status).strip()
        return [text] if text else []

    @staticmethod
    def _is_ok_status(value: str) -> bool:
        return value.strip().lower() in _OK_STATUS_VALUES

    @staticmethod
    def _severity_for_status(field_name: str, value: str) -> str:
        if field_name == "_standardization_failures":
            return "error"
        text = value.lower()
        if any(marker in text for marker in _ERROR_MARKERS):
            return "error"
        return "warning"

    @staticmethod
    def _corrected_field_count(row: dict) -> int:
        count = 0
        for field_name, corrected_value in row.items():
            if not field_name.endswith("_corrected"):
                continue
            base_name = field_name[: -len("_corrected")]
            if base_name not in row:
                continue
            if _normalized_value(corrected_value) != _normalized_value(row.get(base_name)):
                count += 1
        return count


def _normalized_value(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()
