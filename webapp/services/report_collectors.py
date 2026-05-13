"""Helpers that collect processing-report warnings and summaries."""

from __future__ import annotations

from collections import Counter, defaultdict

from webapp.models.processing_report import (
    InvalidDateValue,
    InvalidIdentifierValue,
    MissingInputColumnsBySheet,
    MissingRequiredExportField,
    PerSheetProcessingReport,
    ProcessingReport,
)
from webapp.services.export_schema import EXPORT_MAPPING, canonical_sheet_name, headers_for_sheet
from webapp.services.export_rows import build_row_export_view, resolve_sug_mosad_for_sheet, visible_rows
from webapp.services.report_status_builder import is_invalid_date_component, row_number


# אוסף עמודות קלט חסרות לפי גיליון לצורך דוח סקירת העיבוד.
def collect_missing_input_columns(workbook_dataset) -> list[MissingInputColumnsBySheet]:
    """Collect expected input columns that are missing from each sheet."""
    if workbook_dataset is None:
        return []

    expected_base = [
        "first_name",
        "last_name",
        "father_name",
        "id_number",
        "passport",
        "gender",
    ]

    result = []
    for sheet in workbook_dataset.sheets:
        fields = set(sheet.field_names or [])
        expected = list(expected_base)
        if "birth_date" not in fields:
            expected.extend(["birth_year", "birth_month", "birth_day"])
        if "entry_date" not in fields:
            expected.extend(["entry_year", "entry_month", "entry_day"])

        missing = [field for field in expected if field not in fields]
        if missing:
            result.append(MissingInputColumnsBySheet(sheet_name=sheet.sheet_name, columns=missing))

    return result


# אוסף שדות יצוא חובה שחסרים לאחר סטנדרטיזציה.
def collect_missing_required_export_fields(record) -> list[MissingRequiredExportField]:
    """Collect export fields that are still empty after session injection."""
    if record.workbook_dataset is None:
        return []

    missing = Counter()
    active_mosad_type = record.mosad_types[0] if record.mosad_types else ""

    for sheet_dataset in record.workbook_dataset.sheets:
        export_name = canonical_sheet_name(sheet_dataset.sheet_name)
        schema = headers_for_sheet(export_name)
        data_rows, _ui_cols = visible_rows(sheet_dataset)
        scoped_type = resolve_sug_mosad_for_sheet(
            record.sug_mosad_configs,
            sheet_dataset.sheet_name,
            active_mosad_type,
        )

        for row in data_rows:
            effective_row = build_row_export_view(
                row,
                mosad_id=record.mosad_id or "",
                scoped_sug_mosad=scoped_type,
            )

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


# אוסף ערכי תאריך בעייתיים מתוך סטטוסי השורות.
def collect_invalid_date_values(record) -> list[InvalidDateValue]:
    """Collect invalid birth and entry date components for the report."""
    if record.workbook_dataset is None:
        return []

    invalid = []
    for sheet in record.workbook_dataset.sheets:
        for idx, row in enumerate(sheet.rows):
            row_uid = row.get("_row_uid")
            current_row_number = row_number(sheet, idx)
            for field in ("birth_year", "birth_month", "birth_day", "entry_year", "entry_month", "entry_day"):
                if field not in row:
                    continue
                prefix = "birth" if field.startswith("birth_") else "entry"
                status = row.get(f"{prefix}_date_status") or ""
                corrected = row.get(f"{field}_corrected")
                if not is_invalid_date_component(status, corrected):
                    continue
                invalid.append(
                    InvalidDateValue(
                        sheet_name=sheet.sheet_name,
                        row_number=current_row_number if row_uid is None else None,
                        row_uid=row_uid,
                        source_field=field,
                        raw_value=row.get(field),
                        corrected_value=corrected,
                        status_message=status,
                    )
                )

    return invalid


# אוסף בעיות מזהים אמיתיות לדוח לפי מסנן הסטטוס.
def collect_invalid_identifier_values(record, is_real_identifier_issue) -> list[InvalidIdentifierValue]:
    """Collect identifier rows whose status represents a real problem."""
    if record.workbook_dataset is None:
        return []

    invalid = []
    for sheet in record.workbook_dataset.sheets:
        for idx, row in enumerate(sheet.rows):
            status = row.get("identifier_status") or ""
            if not is_real_identifier_issue(status):
                continue
            row_uid = row.get("_row_uid")
            current_row_number = row_number(sheet, idx)
            for field in ("id_number", "passport"):
                if field not in row:
                    continue
                invalid.append(
                    InvalidIdentifierValue(
                        sheet_name=sheet.sheet_name,
                        row_number=current_row_number if row_uid is None else None,
                        row_uid=row_uid,
                        source_field=field,
                        raw_value=row.get(field),
                        corrected_value=row.get(f"{field}_corrected"),
                        status_message=status,
                    )
                )

    return invalid


# מסכם הודעות פירוט לפי גיליון כדי לצמצם עומס בדוח.
def aggregate_detail_messages_by_sheet(details) -> dict[str, list[tuple[str, int]]]:
    """Group per-sheet validation messages into compact counts."""
    counts = defaultdict(Counter)
    for item in details:
        if item.status_message:
            counts[item.sheet_name][item.status_message] += 1
    return {
        sheet_name: sorted(message_counts.items())
        for sheet_name, message_counts in counts.items()
    }


# מסכם הודעות מזהים לפי גיליון לצורך תצוגת אזהרות.
def aggregate_identifier_messages_by_sheet(details) -> dict[str, list[tuple[str, int]]]:
    """Group identifier validation messages by sheet and count."""
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


# בונה הודעת אזהרה כאשר עמודות חובה ריקות או חסרות בגיליון.
def empty_required_columns_message_for_sheet(
    export_name: str,
    details: list[MissingRequiredExportField],
) -> str:
    """Format one compact warning for missing required export fields."""
    counts_by_field: dict[str, int] = {}
    for item in details:
        if item.sheet_name != export_name or item.rows_affected <= 0:
            continue
        counts_by_field[item.field_name] = counts_by_field.get(item.field_name, 0) + item.rows_affected

    if not counts_by_field:
        return ""

    ordered_fields = []
    for field_name in headers_for_sheet(export_name):
        if field_name in counts_by_field:
            ordered_fields.append(f"{field_name}={counts_by_field[field_name]}")

    return f"עמודות חובה ריקות: {', '.join(ordered_fields)}"


# יוצר אזהרות לפי גיליון עבור דוח העיבוד הסופי.
def build_per_sheet_warnings(
    record,
    rows_exported_by_sheet: dict[str, int],
    report: ProcessingReport,
) -> list[PerSheetProcessingReport]:
    """Build compact per-sheet warnings for the processing report."""
    if record.workbook_dataset is None:
        return []

    missing_input_by_sheet = {item.sheet_name: len(item.columns) for item in report.missing_input_columns}
    invalid_dates_by_sheet = aggregate_detail_messages_by_sheet(report.invalid_date_values or [])
    invalid_ids_by_sheet = aggregate_identifier_messages_by_sheet(report.invalid_identifier_values or [])

    per_sheet = []
    for sheet in record.workbook_dataset.sheets:
        warnings = []
        if missing_input_by_sheet.get(sheet.sheet_name):
            warnings.append(f"{missing_input_by_sheet[sheet.sheet_name]} expected input column(s) missing")
        export_name = canonical_sheet_name(sheet.sheet_name)
        required_columns_message = empty_required_columns_message_for_sheet(
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

