"""Date standardization helpers for the processing pipeline."""

from __future__ import annotations

import logging
from datetime import date as _date, datetime as _dt
from typing import Any, List, Optional, Tuple

from ..data_types import DateFieldType, DateFormatPattern, JsonRow

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

        if year_val is not None and month_val is None and day_val is None:
            main_val_for_engine = year_val
            year_val_for_engine = None
            month_val_for_engine = None
            day_val_for_engine = None
        elif isinstance(year_val, (_dt, _date)):
            main_val_for_engine = year_val
            year_val_for_engine = None
            month_val_for_engine = None
            day_val_for_engine = None
        else:
            main_val_for_engine = None
            year_val_for_engine = year_val
            month_val_for_engine = month_val
            day_val_for_engine = day_val

        try:
            result = pipeline.date_engine.parse_date(
                year_val_for_engine,
                month_val_for_engine,
                day_val_for_engine,
                main_val_for_engine,
                pattern,
                field_type,
            )
            date_result = result

            corrected_year, corrected_month, corrected_day = date_corrected_components(result)
            json_row[f"{year_field}_corrected"] = corrected_year
            json_row[f"{month_field}_corrected"] = corrected_month
            json_row[f"{day_field}_corrected"] = corrected_day
            json_row[f"{prefix}_date_status"] = result.status_text
            json_row[f"_{prefix}_year_auto_completed"] = result.year_was_auto_completed

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

        if date_val is None or date_val == "":
            json_row[f"{year_field}_corrected"] = None
            json_row[f"{month_field}_corrected"] = None
            json_row[f"{day_field}_corrected"] = None
            json_row[f"{prefix}_date_status"] = ""
            json_row[f"_{prefix}_year_auto_completed"] = False
            return failed_fields, date_result

        try:
            result = pipeline.date_engine.parse_date(
                None,
                None,
                None,
                date_val,
                pattern,
                field_type,
            )
            date_result = result

            json_row[f"{year_field}_corrected"] = result.year
            json_row[f"{month_field}_corrected"] = result.month
            json_row[f"{day_field}_corrected"] = result.day
            json_row[f"{prefix}_date_status"] = result.status_text
            json_row[f"_{prefix}_year_auto_completed"] = result.year_was_auto_completed

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
    status = result.status_text or ""

    if status == "ערך תאריך לא תקין":
        return (
            year if year is not None else "",
            month if month is not None else "",
            day if day is not None else "",
        )

    if status == "שנה לא תקינה":
        year = ""
    if status == "חודש לא תקין":
        month = ""
    if status == "יום לא תקין":
        day = ""

    return year, month, day


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
                    new_result = pipeline.date_engine._validate_date(new_yr, mo, dy)
                    new_result.year_was_auto_completed = True
                    new_result = pipeline.date_engine.validate_business_rules(
                        new_result, DateFieldType.BIRTH_DATE
                    )
                    row["birth_year_corrected"] = new_result.year if new_result.year is not None else new_yr
                    row["birth_month_corrected"] = new_result.month if new_result.month is not None else mo
                    row["birth_day_corrected"] = new_result.day if new_result.day is not None else dy
                    row["birth_date_status"] = new_result.status_text
                except Exception:
                    row["birth_year_corrected"] = new_yr

        row.pop("_birth_year_auto_completed", None)
        row.pop("_entry_year_auto_completed", None)
        corrected_rows.append(row)

    return corrected_rows
