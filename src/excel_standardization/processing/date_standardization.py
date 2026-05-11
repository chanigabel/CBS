"""Date standardization helpers for the processing pipeline."""

from __future__ import annotations

import logging
from typing import Any, List, Optional, Tuple

from ..data_types import DateFieldType, DateFormatPattern, DateInput, JsonRow

logger = logging.getLogger(__name__)


# הפונקציה מזהה ברמת גיליון אם התאריכים נראים כ־DD/MM או MM/DD לפני parsing.
def detect_date_format_pattern(rows: List[JsonRow]) -> DateFormatPattern:
    """Detect whether date values in this dataset use DDMM or MMDD ordering."""
    date_fields = (
        "birth_date",
        "entry_date",
        "birth_year",
        "entry_year",
    )
    ddmm = 0
    mmdd = 0

    for row in rows[:20]:
        for field in date_fields:
            val = row.get(field)
            if not val or not isinstance(val, str):
                continue
            s = val.replace(".", "/")
            if "/" not in s:
                continue
            parts = s.split("/")
            if len(parts) < 2:
                continue
            try:
                a, b = int(parts[0]), int(parts[1])
                if a > 12 and b <= 12:
                    ddmm += 1
                elif b > 12 and a <= 12:
                    mmdd += 1
            except (ValueError, TypeError):
                pass

    return DateFormatPattern.MMDD if mmdd > ddmm else DateFormatPattern.DDMM


# הפונקציה מפעילה נרמול תאריכי לידה וכניסה עבור שורת Dataset אחת.
def apply_date_standardization(
    pipeline: Any,
    json_row: JsonRow,
    row_number: Optional[int] = None,
) -> List[str]:
    """Apply DateEngine to date fields in the row."""
    failed_fields: List[str] = []

    failures, birth_result = normalize_date_field(
        pipeline,
        json_row,
        "birth",
        DateFieldType.BIRTH_DATE,
        row_number,
    )
    failed_fields.extend(failures)

    failures, entry_result = normalize_date_field(
        pipeline,
        json_row,
        "entry",
        DateFieldType.ENTRY_DATE,
        row_number,
    )
    failed_fields.extend(failures)

    if birth_result is not None and entry_result is not None:
        try:
            if not pipeline.date_engine.validate_entry_before_birth(birth_result, entry_result):
                warning = "תאריך כניסה לפני תאריך לידה"
                existing_status = json_row.get("entry_date_status", "")
                if existing_status:
                    json_row["entry_date_status"] = f"{existing_status} | {warning}"
                else:
                    json_row["entry_date_status"] = warning
        except Exception:
            pass

    return failed_fields


# הפונקציה מנרמלת שדה תאריך יחיד או תאריך מפוצל לשדות corrected עקביים.
def normalize_date_field(
    pipeline: Any,
    json_row: JsonRow,
    prefix: str,
    field_type: DateFieldType,
    row_number: Optional[int] = None,
):
    """Normalize one date field group (birth or entry)."""
    failed_fields: List[str] = []
    date_result = None

    pattern = getattr(pipeline, "_date_format_pattern", DateFormatPattern.DDMM)

    year_field = f"{prefix}_year"
    month_field = f"{prefix}_month"
    day_field = f"{prefix}_day"
    date_field = f"{prefix}_date"

    has_split = year_field in json_row or month_field in json_row or day_field in json_row
    has_single = date_field in json_row

    if has_split:
        year_val = json_row.get(year_field)
        month_val = json_row.get(month_field)
        day_val = json_row.get(day_field)

        try:
            result = pipeline.date_engine.parse_input(
                DateInput(
                    source_kind="split",
                    field_type=field_type,
                    raw_year=year_val,
                    raw_month=month_val,
                    raw_day=day_val,
                    pattern=pattern,
                    reference_date=getattr(pipeline, "_reference_date", None),
                )
            )
            date_result = result

            corrected_year, corrected_month, corrected_day = date_corrected_components(result)
            json_row[f"{year_field}_corrected"] = corrected_year
            json_row[f"{month_field}_corrected"] = corrected_month
            json_row[f"{day_field}_corrected"] = corrected_day
            json_row[f"{prefix}_date_status"] = result.status_text
            json_row[f"_{prefix}_year_auto_completed"] = result.year_was_auto_completed
            if result.original_year_value is not None:
                json_row[f"_{prefix}_year_original_two_digit"] = result.original_year_value
                json_row[f"_{prefix}_year_reference_year"] = result.reference_year

        except Exception as e:
            json_row[f"{year_field}_corrected"] = ""
            json_row[f"{month_field}_corrected"] = ""
            json_row[f"{day_field}_corrected"] = ""
            json_row[f"{prefix}_date_status"] = "ערך תאריך לא תקין"
            json_row[f"_{prefix}_year_auto_completed"] = False
            failed_fields.extend([year_field, month_field, day_field])

            row_info = f"row {row_number}" if row_number is not None else "unknown row"
            logger.error(
                f"Date standardization failed for split date fields '{prefix}_*' at {row_info}: {str(e)}. "
                f"Original values: year={year_val}, month={month_val}, day={day_val}"
            )

    elif has_single:
        date_val = json_row.get(date_field)
        year_field = f"{prefix}_year"
        month_field = f"{prefix}_month"
        day_field = f"{prefix}_day"
        source_serial_field = f"_{date_field}_source_is_excel_date_serial"
        legacy_source_serial_field = f"_{prefix}_source_is_excel_date_serial"

        if date_val is None or date_val == "":
            json_row[f"{year_field}_corrected"] = None
            json_row[f"{month_field}_corrected"] = None
            json_row[f"{day_field}_corrected"] = None
            json_row[f"{prefix}_date_status"] = ""
            json_row[f"_{prefix}_year_auto_completed"] = False
            return failed_fields, date_result

        try:
            result = pipeline.date_engine.parse_input(
                DateInput(
                    source_kind="single",
                    field_type=field_type,
                    raw_value=date_val,
                    pattern=pattern,
                    reference_date=getattr(pipeline, "_reference_date", None),
                    source_is_excel_date_serial=bool(
                        json_row.get(source_serial_field, False)
                        or json_row.get(legacy_source_serial_field, False)
                    ),
                )
            )
            date_result = result

            corrected_year, corrected_month, corrected_day = date_corrected_components(result)
            json_row[f"{year_field}_corrected"] = corrected_year
            json_row[f"{month_field}_corrected"] = corrected_month
            json_row[f"{day_field}_corrected"] = corrected_day
            json_row[f"{prefix}_date_status"] = result.status_text
            json_row[f"_{prefix}_year_auto_completed"] = result.year_was_auto_completed
            if result.original_year_value is not None:
                json_row[f"_{prefix}_year_original_two_digit"] = result.original_year_value
                json_row[f"_{prefix}_year_reference_year"] = result.reference_year

        except Exception as e:
            json_row[f"{year_field}_corrected"] = None
            json_row[f"{month_field}_corrected"] = None
            json_row[f"{day_field}_corrected"] = None
            json_row[f"{prefix}_date_status"] = ""
            json_row[f"_{prefix}_year_auto_completed"] = False
            failed_fields.append(date_field)

            row_info = f"row {row_number}" if row_number is not None else "unknown row"
            logger.error(
                f"Date standardization failed for field '{date_field}' at {row_info}: {str(e)}. "
                f"Original value: '{date_val}'"
            )

    return failed_fields, date_result


# הפונקציה ממירה תוצאת parsing לרכיבי שנה/חודש/יום בטוחים ל־UI וליצוא.
def date_corrected_components(result) -> Tuple[Any, Any, Any]:
    """Return UI/export-safe corrected date components."""
    year = result.year
    month = result.month
    day = result.day
    invalid_components = set(getattr(result, "invalid_components", []) or [])
    missing_components = set(getattr(result, "missing_components", []) or [])
    status_code = getattr(result, "status_code", "") or ""

    if "year" in invalid_components or "year" in missing_components:
        year = ""
    if "month" in invalid_components or "month" in missing_components:
        month = ""
    if "day" in invalid_components or "day" in missing_components:
        day = ""

    if status_code in {
        "empty_cell",
        "empty_optional_entry",
        "unrecognized_format",
        "invalid_format",
        "invalid_length",
        "unclear_date",
        "unparseable",
        # Impossible / out-of-range dates must never populate corrected fields.
        # These are business-domain rejections where the entire date is meaningless.
        "impossible_year",
        "future_birth",
        "future_entry",
        "late_entry",
        "year_before_1906",
        "unrecognized_numeric_date",
    }:
        # Force-blank ALL components — these status codes mean the date is
        # invalid or impossible and must not appear in corrected output fields.
        return ("", "", "")

    if (
        getattr(result, "source_kind", "") == "single_numeric"
        and not getattr(result, "is_calendar_valid", False)
    ):
        return ("", "", "")

    return (
        year if year is not None else "",
        month if month is not None else "",
        day if day is not None else "",
    )


# הפונקציה מתקנת רק שנות לידה דו־ספרתיות לפי רוב הגיליון ומסירה תגיות פנימיות.
def apply_birth_year_majority_correction(pipeline: Any, rows: List[JsonRow]) -> List[JsonRow]:
    """One-way list-level majority correction for birth years in the web/JSON path."""

    from ..data_types import DateFieldType

    def _get_corrected_year(row):
        if "birth_year_corrected" in row:
            try:
                return int(row["birth_year_corrected"])
            except (TypeError, ValueError):
                return None
        return None

    auto_1900s = sum(
        1 for r in rows
        if r.get("_birth_year_auto_completed") is True
        and _get_corrected_year(r) is not None
        and 1900 <= _get_corrected_year(r) <= 1999
    )
    auto_2000s = sum(
        1 for r in rows
        if r.get("_birth_year_auto_completed") is True
        and _get_corrected_year(r) is not None
        and 2000 <= _get_corrected_year(r) <= 2099
    )

    total_auto = auto_1900s + auto_2000s
    do_correction = total_auto > 0 and auto_1900s > auto_2000s

    corrected_rows = []
    for row in rows:
        row = dict(row)
        is_auto = row.get("_birth_year_auto_completed") is True
        yr = _get_corrected_year(row)

        if do_correction and is_auto and yr is not None and 2000 <= yr <= 2099:
            new_yr = yr - 100

            if "birth_year_corrected" in row:
                mo = row.get("birth_month_corrected")
                dy = row.get("birth_day_corrected")
                try:
                    new_result = pipeline.date_engine._validate_date(
                        new_yr,
                        mo,
                        dy,
                        source_kind="majority_correction",
                        reference_date=getattr(pipeline, "_reference_date", None),
                    )
                    new_result.year_was_auto_completed = True
                    new_result.original_year_value = row.get("_birth_year_original_two_digit")
                    new_result.reference_year = row.get("_birth_year_reference_year")
                    new_result = pipeline.date_engine.validate_business_rules(
                        new_result,
                        DateFieldType.BIRTH_DATE,
                        reference_date=getattr(pipeline, "_reference_date", None),
                    )
                    cy, cm, cd = date_corrected_components(new_result)
                    row["birth_year_corrected"] = cy if cy != "" else new_yr
                    row["birth_month_corrected"] = cm
                    row["birth_day_corrected"] = cd
                    row["birth_date_status"] = new_result.status_text
                    row["_birth_year_majority_corrected"] = True
                except Exception:
                    row["birth_year_corrected"] = new_yr

        row.pop("_birth_year_auto_completed", None)
        row.pop("_entry_year_auto_completed", None)
        row.pop("_birth_year_original_two_digit", None)
        row.pop("_entry_year_original_two_digit", None)
        row.pop("_birth_year_reference_year", None)
        row.pop("_entry_year_reference_year", None)
        corrected_rows.append(row)

    return corrected_rows
