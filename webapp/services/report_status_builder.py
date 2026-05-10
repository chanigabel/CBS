"""Helpers for compact processing-report status and reason text."""

from __future__ import annotations

from webapp.models.processing_report import ProcessingReport

_STAGE_ORDER = ["upload", "extract", "standardize", "validate", "export"]


# קובע האם רכיב תאריך צריך להיחשב כבעייתי בדוח העיבוד.
def is_invalid_date_component(status: str, corrected) -> bool:
    """Return True when a date status indicates an invalid corrected value."""
    if not status:
        return False
    invalid_status = (
        status == "ערך תאריך לא תקין"
        or "לא תקין" in status
        or "לא תקינה" in status
        or "לא קיים" in status
    )
    return invalid_status and (corrected is None or str(corrected).strip() == "")


# מסנן סטטוס מזהה כדי לדווח רק על בעיות אמיתיות למשתמש.
def is_real_identifier_issue(status: str) -> bool:
    return bool(status) and ("חסר מזהים" in status or "לא תקינה" in status)


# מחשב מספר שורה אמיתי ב־Excel עבור הודעות דוח.
def row_number(sheet, idx: int) -> int:
    """Return the 1-based workbook row number for a sheet row index."""
    return int(sheet.header_row or 1) + int(sheet.header_rows_count or 1) + idx + 1


# מעדכן סטטוס כולל של דוח העיבוד לפי שלבים, אזהרות ושגיאות.
def refresh_status(report: ProcessingReport) -> None:
    """Recompute processing status and its human-readable reason."""
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
    report.status_reason = status_reason(report)


# מחזיר טקסט קצר שמסביר את הסטטוס הנוכחי של הדוח.
def status_reason(report: ProcessingReport) -> str:
    """Build a compact reason string for the current report state."""
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

