"""Build user-facing workbook processing reports from session state."""

from __future__ import annotations

from collections import Counter
from typing import Any, Iterable

from webapp.models.report import (
    ManualEditsSummary,
    ReportIssueGroup,
    ReportSummary,
    ReportIssue,
    SheetReport,
    WorkbookProcessingReport,
)
from webapp.services.report_state import source_row_count
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
            display_status=self._workbook_display_status(record.status, record.workbook_dataset),
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
        total_source_rows = 0
        total_current_rows = 0
        total_warning_rows = 0
        total_error_rows = 0
        total_corrected_fields = 0
        total_manual_rows = 0
        sheets: list[SheetReport] = []

        for sheet in workbook_dataset.sheets:
            sheet_manual_rows = self._manual_rows_for_sheet(record.edits, sheet.sheet_name)
            sheet_report, sheet_issues = self._sheet_report(
                sheet,
                include_details=include_details,
                manual_row_count=sheet_manual_rows,
                standardized=record.status == "standardized",
            )
            sheets.append(sheet_report)
            if include_details:
                issues.extend(sheet_issues)
            total_source_rows += sheet_report.source_row_count
            total_current_rows += sheet_report.current_row_count
            total_warning_rows += sheet_report.rows_with_warnings
            total_error_rows += sheet_report.rows_with_errors
            total_corrected_fields += sheet_report.corrected_fields
            total_manual_rows += sheet_manual_rows

        report.sheets = sheets
        report.summary = ReportSummary(
            total_sheets=len(sheets),
            total_rows=total_source_rows,
            source_rows=total_source_rows,
            current_rows=total_current_rows,
            rows_deleted=max(total_source_rows - total_current_rows, 0),
            rows_processed=total_current_rows if record.status == "standardized" else 0,
            rows_changed_automatically=total_corrected_fields,
            rows_changed_manually=total_manual_rows,
            manual_edit_rows=total_manual_rows,
            manual_edit_actions=report.manual_edits.edited_actions,
            edited_cells=report.manual_edits.edited_cells,
            rows_with_warnings=total_warning_rows,
            rows_with_errors=total_error_rows,
            rows_without_issues=max(total_current_rows - total_warning_rows - total_error_rows, 0),
            corrected_fields=total_corrected_fields,
        )
        if record.status == "standardized":
            report.display_status = (
                "בוצע עם אזהרות"
                if total_warning_rows or total_error_rows
                else "בוצע"
            )
        if include_details:
            report.issues = issues
        return report

    def _sheet_report(
        self,
        sheet,
        include_details: bool = False,
        manual_row_count: int = 0,
        standardized: bool = False,
    ) -> tuple[SheetReport, list[ReportIssue]]:
        status_counts: dict[str, Counter[str]] = {}
        warning_rows = 0
        error_rows = 0
        corrected_fields = 0
        issue_count = 0
        issues: list[ReportIssue] = []
        issue_groups: dict[tuple[str, str], dict[str, Any]] = {}

        for row_number, row in enumerate(sheet.rows, start=1):
            has_warning = False
            has_error = False
            seen_row_messages: set[tuple[str, str]] = set()

            for field_name, raw_status in row.items():
                if not self._is_status_field(field_name):
                    continue
                values = self._status_values(raw_status)
                for value in values:
                    if self._is_ok_status(value):
                        continue
                    status_counts.setdefault(field_name, Counter())[value] += 1
                    severity = self._severity_for_status(field_name, value)
                    group_key = (severity, value)
                    if group_key not in seen_row_messages:
                        group = issue_groups.setdefault(
                            group_key,
                            {
                                "label": value,
                                "severity": severity,
                                "count": 0,
                                "row_numbers": [],
                                "row_uids": [],
                                "field_names": [],
                            },
                        )
                        group["count"] += 1
                        group["row_numbers"].append(row_number)
                        row_uid = str(row.get("_row_uid") or "")
                        if row_uid and row_uid not in group["row_uids"]:
                            group["row_uids"].append(row_uid)
                        if field_name not in group["field_names"]:
                            group["field_names"].append(field_name)
                        seen_row_messages.add(group_key)
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

        source_rows = source_row_count(sheet)
        current_rows = len(sheet.rows)

        return SheetReport(
            sheet_name=sheet.sheet_name,
            status=self._sheet_display_status(
                standardized=standardized,
                rows_with_warnings=warning_rows,
                rows_with_errors=error_rows,
            ),
            source_row_count=source_rows,
            current_row_count=current_rows,
            row_count=current_rows,
            rows_processed=current_rows if standardized else 0,
            rows_changed_automatically=corrected_fields,
            rows_changed_manually=manual_row_count,
            rows_deleted=max(source_rows - current_rows, 0),
            column_count=len(sheet.field_names or []),
            rows_with_warnings=warning_rows,
            rows_with_errors=error_rows,
            corrected_fields=corrected_fields,
            issues_count=issue_count,
            status_counts={
                field: dict(sorted(counter.items()))
                for field, counter in sorted(status_counts.items())
            },
            issue_groups=[
                ReportIssueGroup(
                    label=str(group["label"]),
                    severity=str(group["severity"]),
                    count=int(group["count"]),
                    row_numbers=sorted(dict.fromkeys(int(v) for v in group["row_numbers"])),
                    row_uids=sorted(dict.fromkeys(str(v) for v in group["row_uids"])),
                    field_names=sorted(dict.fromkeys(str(v) for v in group["field_names"])),
                )
                for group in sorted(issue_groups.values(), key=lambda item: (item["severity"], item["label"]))
            ],
        ), issues

    @staticmethod
    def _manual_rows_for_sheet(edits: dict, sheet_name: str) -> int:
        row_uids = set()
        for key in edits:
            if not isinstance(key, tuple) or len(key) != 3:
                continue
            edit_sheet, row_uid, _field_name = key
            if str(edit_sheet) == str(sheet_name):
                row_uids.add(str(row_uid))
        return len(row_uids)

    @staticmethod
    def _sheet_display_status(
        *,
        standardized: bool,
        rows_with_warnings: int,
        rows_with_errors: int,
    ) -> str:
        if rows_with_errors:
            return "נכשל"
        if not standardized:
            return "ממתין לעיבוד"
        if rows_with_warnings:
            return "בוצע עם אזהרות"
        return "בוצע"

    @staticmethod
    def _workbook_display_status(status: str, workbook_dataset) -> str:
        if status != "standardized":
            return "ממתין לעיבוד"
        if workbook_dataset is None:
            return "ממתין לעיבוד"
        return "בוצע"

    @staticmethod
    def _manual_edits_summary(edits: dict) -> ManualEditsSummary:
        sheets = set()
        fields = set()
        rows = set()
        for key in edits:
            if not isinstance(key, tuple) or len(key) != 3:
                continue
            sheet_name, _row_uid, field_name = key
            sheets.add(str(sheet_name))
            fields.add(str(field_name))
            rows.add((str(sheet_name), str(_row_uid)))
        return ManualEditsSummary(
            edited_cells=len(edits),
            edited_rows=len(rows),
            edited_actions=len(edits),
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
