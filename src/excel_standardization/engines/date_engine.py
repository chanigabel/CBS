"""Date parsing and validation rules for birth and entry dates.

The engine parses split or combined date values, expands two-digit years,
validates calendar dates, and applies the birth-date and entry-date business
rules used by the active pipeline.
"""

from datetime import date, datetime
import logging
import re
from typing import Optional

from ..data_types import DateInput, DateParseResult, DateFormatPattern, DateFieldType


logger = logging.getLogger(__name__)


STATUS_EMPTY_CELL = "תא ריק"
STATUS_INVALID_DATE_VALUE = "ערך תאריך לא תקין"
STATUS_UNPARSEABLE = "תוכן לא ניתן לפריקה"
STATUS_INVALID_DAY = "יום לא תקין"
STATUS_INVALID_MONTH = "חודש לא תקין"
STATUS_INVALID_YEAR = "שנה לא תקינה"
STATUS_DATE_NOT_EXISTS = "תאריך לא קיים"
STATUS_BEFORE_1906 = "שנה לפני 1906"
STATUS_LATE_ENTRY = "תאריך כניסה מאוחר מהתאריך שנקבע"
STATUS_FUTURE_BIRTH = "תאריך לידה עתידי"
STATUS_FUTURE_ENTRY = "תאריך כניסה עתידי"
STATUS_MISSING_MONTH_DAY = "חסר חודש ויום"
STATUS_INVALID_LENGTH = "אורך תאריך לא תקין"
STATUS_UNCLEAR_DATE = "תאריך לא ברור"
STATUS_INVALID_FORMAT = "פורמט תאריך לא תקין"
STATUS_UNRECOGNIZED_FORMAT = "פורמט תאריך לא מזוהה"
STATUS_NO_SEPARATOR = "אין מפריד בתאריך"
STATUS_MISSING_YEAR = "חסר שנה"
STATUS_MISSING_MONTH = "חסר חודש"
STATUS_MISSING_DAY = "חסר יום"
STATUS_MISSING_YEAR_DEFAULTED = "שנה חסרה והושלמה"
STATUS_EXCEL_SERIAL_PARSED = "פורק מתאריך סידורי"
STATUS_NUMERIC_DATE_UNRECOGNIZED = "מספר לא הוכר כתאריך"
STATUS_EXCEL_SERIAL_NOT_RECOGNIZED = STATUS_NUMERIC_DATE_UNRECOGNIZED
STATUS_AMBIGUOUS_NUMERIC_DATE = STATUS_UNCLEAR_DATE
STATUS_TRAILING_TEXT_IGNORED = "טקסט נוסף הוסר מהתאריך"
STATUS_IMPOSSIBLE_YEAR = "שנה לא סבירה"
STATUS_SPLIT_FULL_DATE_CONFLICT = "ערכים סותרים בעמודות תאריך מפוצלות"
STATUS_SPLIT_FULL_DATE_FROM_DAY = "תאריך מלא פורק מעמודת יום"
STATUS_SPLIT_FULL_DATE_FROM_MONTH = "תאריך מלא פורק מעמודת חודש"
STATUS_SPLIT_FULL_DATE_FROM_YEAR = "תאריך מלא פורק מעמודת שנה"


# המנוע אחראי לפענוח תאריכים, תיקון רכיבים והרצת חוקי תאריך.
class DateEngine:
    def __init__(self, reference_date: Optional[date] = None) -> None:
        self.reference_date = reference_date or date.today()

    # הפונקציה יוצרת תוצאת תאריך ריקה כאשר אין ערך קלט לעיבוד.
    def _blank_result(
        self,
        *,
        source_kind: str = "",
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        ref = self._get_reference_date(reference_date)
        return DateParseResult(
            year=None,
            month=None,
            day=None,
            is_valid=False,
            status_text="",
            severity="error",
            status_code="",
            reference_year=ref.year,
            is_calendar_valid=False,
            is_business_valid=False,
            source_kind=source_kind,
        )

    # ----------------------------------------------------
    # MAIN ENTRY
    # ----------------------------------------------------

    # הפונקציה משמשת נקודת כניסה לנרמול תאריך משדה יחיד או מעמודות שנה/חודש/יום.
    def parse_date(
        self,
        year_val,
        month_val,
        day_val,
        main_val,
        pattern: DateFormatPattern,
        field_type: DateFieldType,
        source_kind: Optional[str] = None,
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        ref = self._get_reference_date(reference_date)

        if source_kind is None:
            if (
                self._has_any_split_component(year_val, month_val, day_val)
                and not (
                    main_val is not None
                    and self._is_empty(month_val)
                    and self._is_empty(day_val)
                )
            ):
                source_kind = "split"
            elif self._is_empty(main_val):
                source_kind = "missing"
            else:
                source_kind = "single"

        date_input = DateInput(
            source_kind=source_kind,
            field_type=field_type,
            raw_value=main_val,
            raw_year=year_val,
            raw_month=month_val,
            raw_day=day_val,
            pattern=pattern,
            reference_date=ref,
            source_is_excel_date_serial=False,
        )
        return self.parse_input(date_input)

    # הפונקציה מפעילה parsing לפי מודל קלט מובנה ושומרת על תאימות לממשק הישן.
    def parse_input(self, date_input: DateInput) -> DateParseResult:
        pattern = date_input.pattern or DateFormatPattern.DDMM
        ref = self._get_reference_date(date_input.reference_date)

        if date_input.source_kind == "split":
            result = self.parse_from_split_columns(
                date_input.raw_year,
                date_input.raw_month,
                date_input.raw_day,
                reference_date=ref,
            )
        elif (
            date_input.source_kind == "single"
            and isinstance(date_input.raw_value, (int, float))
            and not isinstance(date_input.raw_value, bool)
        ):
            if date_input.source_is_excel_date_serial:
                result = self._parse_excel_serial_date(date_input.raw_value, reference_date=ref)
            else:
                result = self.parse_from_main_value(
                    date_input.raw_value,
                    pattern,
                    reference_date=ref,
                )
        elif date_input.source_kind == "missing":
            result = self._blank_result(source_kind="missing", reference_date=ref)
            result.status_text = STATUS_EMPTY_CELL
            result.status_code = "empty_cell"
        else:
            result = self.parse_from_main_value(
                date_input.raw_value,
                pattern,
                reference_date=ref,
            )

        if (
            date_input.source_is_excel_date_serial
            and not (
                isinstance(date_input.raw_value, (int, float))
                and not isinstance(date_input.raw_value, bool)
            )
            and result.is_calendar_valid
        ):
            self._set_status_preserving_existing(result, STATUS_EXCEL_SERIAL_PARSED)
            if result.status_code == "ok":
                result.status_code = "excel_serial_parsed"
            if result.severity == "ok":
                result.severity = "warning"

        if not result.source_kind:
            result.source_kind = date_input.source_kind
        return self.validate_business_rules(result, date_input.field_type, reference_date=ref)

    # ----------------------------------------------------
    # SPLIT COLUMNS
    # ----------------------------------------------------

    # הפונקציה מפענחת תאריך שמגיע משלוש עמודות נפרדות בגיליון.
    def parse_from_split_columns(
        self,
        year_val,
        month_val,
        day_val,
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        ref = self._get_reference_date(reference_date)
        result = self._blank_result(source_kind="split", reference_date=ref)

        split_values = {"year": year_val, "month": month_val, "day": day_val}
        full_date_columns = [
            column
            for column, value in split_values.items()
            if not self._is_empty(value) and self._looks_like_full_date_value(value)
        ]
        if full_date_columns:
            if len(full_date_columns) == 1:
                source_column = full_date_columns[0]
                other_values = [
                    (None if column == "year" and self._is_split_year_zero(value) else value)
                    for column, value in split_values.items()
                    if column != source_column
                ]
                if all(self._is_empty(value) for value in other_values):
                    parsed = self.parse_date_value(
                        split_values[source_column],
                        DateFormatPattern.DDMM,
                        reference_date=ref,
                    )
                    parsed.source_kind = "split"
                    parsed.status_code = f"full_date_from_{source_column}_column"
                    parsed.status_text = self._split_full_date_status(source_column)
                    return parsed

            result.status_text = STATUS_SPLIT_FULL_DATE_CONFLICT
            result.status_code = "split_full_date_conflict"
            return result

        yr, year_ok, year_missing = self._coerce_split_component(year_val, zero_is_missing=True)
        mo, month_ok, month_missing = self._coerce_split_component(month_val)
        dy, day_ok, day_missing = self._coerce_split_component(day_val)

        result.year = yr
        result.month = mo
        result.day = dy

        if year_missing:
            result.missing_components.append("year")
        elif not year_ok:
            result.invalid_components.append("year")

        if month_missing:
            result.missing_components.append("month")
        elif not month_ok:
            result.invalid_components.append("month")

        if day_missing:
            result.missing_components.append("day")
        elif not day_ok:
            result.invalid_components.append("day")

        if result.invalid_components:
            result.status_text = STATUS_INVALID_DATE_VALUE
            result.status_code = "invalid_split_component"
            return result

        if result.missing_components:
            result.status_text = self._missing_components_status(result.missing_components)
            result.status_code = "missing_" + "_".join(result.missing_components)
            return result

        assert yr is not None
        assert mo is not None
        assert dy is not None

        auto_year: Optional[int] = None
        if 0 <= yr < 100:
            auto_year = yr
            result.year_was_auto_completed = True
            result.original_year_value = yr
            result.original_year_digits = 1 if yr < 10 else 2
            result.reference_year = ref.year
            yr = self._expand_two_digit_year(yr, reference_date=ref)

        result = self._validate_date(yr, mo, dy, source_kind="split", reference_date=ref)
        if auto_year is not None:
            self._mark_auto_completed_year(result, auto_year, ref)
        return result

    # ----------------------------------------------------
    # MAIN VALUE
    # ----------------------------------------------------

    # הפונקציה מפענחת תאריך שמגיע מערך יחיד ושומרת תאימות לחתימות קיימות.
    def parse_from_main_value(
        self,
        raw_value,
        pattern: DateFormatPattern,
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        return self.parse_date_value(raw_value, pattern, reference_date=reference_date)

    # הפונקציה בוחרת אסטרטגיית parsing לפי סוג הערך ודפוס התאריך שזוהה.
    def parse_date_value(
        self,
        raw_value,
        pattern: DateFormatPattern,
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        ref = self._get_reference_date(reference_date)
        result = self._blank_result(source_kind="single", reference_date=ref)

        if raw_value is None:
            result.status_text = STATUS_EMPTY_CELL
            result.status_code = "empty_cell"
            return result

        txt = str(raw_value).strip()
        if txt == "":
            result.status_text = STATUS_EMPTY_CELL
            result.status_code = "empty_cell"
            return result

        if isinstance(raw_value, (datetime, date)):
            dt = raw_value.date() if isinstance(raw_value, datetime) else raw_value
            result.year = dt.year
            result.month = dt.month
            result.day = dt.day
            result.is_valid = True
            result.is_calendar_valid = True
            result.is_business_valid = True
            result.severity = "ok"
            result.status_code = "ok"
            return result

        if isinstance(raw_value, float) and raw_value.is_integer():
            raw_value = int(raw_value)
            txt = str(raw_value)

        if self._contains_month_name(txt):
            return self._parse_mixed_month_numeric(txt, reference_date=ref)

        if txt.isdigit():
            return self._parse_numeric_date_string(txt, reference_date=ref)

        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", txt)
        if m:
            try:
                yr = int(m.group(1))
                mo = int(m.group(2))
                dy = int(m.group(3))
                return self._validate_date(yr, mo, dy, source_kind="single", reference_date=ref)
            except Exception:
                pass

        if ("/" in txt or "." in txt or "-" in txt) and any(ch.isdigit() for ch in txt):
            txt2 = self._normalize_date_separators(txt)
            return self._parse_separated_date_string(txt2, pattern, reference_date=ref)

        result.status_text = STATUS_UNRECOGNIZED_FORMAT
        result.status_code = "unrecognized_format"
        return result

    # ------------------------------------------------------------------
    # Public compatibility wrappers (used by unit tests / legacy callers)
    # ------------------------------------------------------------------

    # הפונקציה מרחיבה שנה דו־ספרתית למאה המתאימה לפי כללי המערכת.
    def expand_two_digit_year(self, year: int) -> int:
        return self._expand_two_digit_year(year)

    # הפונקציה מפענחת מחרוזת תאריך ספרתית ללא מפרידים.
    def parse_numeric_date_string(self, txt: str) -> DateParseResult:
        if txt is None:
            r = self._blank_result(source_kind="single")
            r.status_text = STATUS_INVALID_FORMAT
            r.status_code = "invalid_format"
            return r
        s = str(txt).strip()
        if not s.isdigit():
            r = self._blank_result(source_kind="single")
            r.status_text = STATUS_INVALID_FORMAT
            r.status_code = "invalid_format"
            return r
        return self._parse_numeric_date_string(s)

    # הפונקציה מפענחת מחרוזת תאריך עם מפרידים לפי DD/MM או MM/DD.
    def parse_separated_date_string(self, txt: str, pattern: DateFormatPattern) -> DateParseResult:
        if txt is None:
            r = self._blank_result(source_kind="single")
            r.status_text = STATUS_NO_SEPARATOR
            r.status_code = "no_separator"
            return r
        s = str(txt).strip()
        if "/" not in s and "." not in s and "-" not in s:
            r = self._blank_result(source_kind="single")
            r.status_text = STATUS_NO_SEPARATOR
            r.status_code = "no_separator"
            return r
        s2 = self._normalize_date_separators(s)
        return self._parse_separated_date_string(s2, pattern)

    # הפונקציה מחשבת גיל לצורך בדיקות עסקיות של תאריך לידה וכניסה.
    def calculate_age(self, *args) -> int:
        if len(args) == 2 and isinstance(args[0], date) and isinstance(args[1], date):
            return self._calculate_age(args[0], args[1])
        if len(args) == 3:
            birth = date(int(args[0]), int(args[1]), int(args[2]))
            return self._calculate_age(birth, self._get_reference_date(None))
        raise TypeError("calculate_age expects (birth, today) or (year, month, day)")

    # ----------------------------------------------------
    # NUMERIC DATE
    # ----------------------------------------------------

    # הפונקציה מיישמת את parsing הספרות הפנימי ומחזירה רכיבי תאריך.
    def _parse_numeric_date_string(
        self,
        txt: str,
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        ref = self._get_reference_date(reference_date)
        result = self._blank_result(source_kind="single_numeric", reference_date=ref)

        try:
            auto_year: Optional[int] = None

            if len(txt) == 8:
                ddmm = self._validate_date(
                    int(txt[4:8]),
                    int(txt[2:4]),
                    int(txt[0:2]),
                    source_kind="single_numeric",
                    reference_date=ref,
                )
                if ddmm.is_calendar_valid:
                    return ddmm

                mmdd = self._validate_date(
                    int(txt[4:8]),
                    int(txt[0:2]),
                    int(txt[2:4]),
                    source_kind="single_numeric",
                    reference_date=ref,
                )
                if mmdd.is_calendar_valid:
                    return mmdd

                if ddmm.status_code == "date_not_exists":
                    return ddmm
                if mmdd.status_code == "date_not_exists":
                    return mmdd

                return self._invalid_numeric_result(ddmm.status_text, ddmm.status_code, ref)

            elif len(txt) == 6:
                dy = int(txt[0:2])
                mo = int(txt[2:4])
                auto_year = int(txt[4:6])
                yr = self._expand_two_digit_year(auto_year, reference_date=ref)
                parsed = self._validate_date(yr, mo, dy, source_kind="single_numeric", reference_date=ref)
                self._mark_auto_completed_year(parsed, auto_year, ref)
                if parsed.is_calendar_valid:
                    return parsed

                fallback_dy = int(txt[0:1])
                fallback_mo = int(txt[1:2])
                fallback_yr = int(txt[2:6])
                fallback = self._validate_date(
                    fallback_yr,
                    fallback_mo,
                    fallback_dy,
                    source_kind="single_numeric",
                    reference_date=ref,
                )
                if fallback.is_calendar_valid:
                    return fallback

                mmdd_auto_year = int(txt[4:6])
                mmdd_yr = self._expand_two_digit_year(mmdd_auto_year, reference_date=ref)
                mmdd = self._validate_date(
                    mmdd_yr,
                    int(txt[0:2]),
                    int(txt[2:4]),
                    source_kind="single_numeric",
                    reference_date=ref,
                )
                self._mark_auto_completed_year(mmdd, mmdd_auto_year, ref)
                if mmdd.is_calendar_valid:
                    return mmdd

                for candidate in (parsed, fallback, mmdd):
                    if candidate.status_code == "date_not_exists":
                        return candidate
                return self._invalid_numeric_result(parsed.status_text, parsed.status_code, ref)

            elif len(txt) == 4:
                yr_int = int(txt)
                if 1900 <= yr_int <= 2100:
                    result.year = yr_int
                    result.month = None
                    result.day = None
                    result.is_valid = False
                    result.status_text = STATUS_MISSING_MONTH_DAY
                    result.status_code = "missing_month_day"
                    result.missing_components = ["month", "day"]
                    return result

                dy = int(txt[0:1])
                mo = int(txt[1:2])
                auto_year = int(txt[2:4])
                yr = self._expand_two_digit_year(auto_year, reference_date=ref)
                parsed = self._validate_date(yr, mo, dy, source_kind="single_numeric", reference_date=ref)
                self._mark_auto_completed_year(parsed, auto_year, ref)
                if parsed.is_calendar_valid:
                    return parsed

                fallback_auto_year = int(txt[2:4])
                fallback_yr = self._expand_two_digit_year(fallback_auto_year, reference_date=ref)
                fallback = self._validate_date(
                    fallback_yr,
                    int(txt[0:1]),
                    int(txt[1:2]),
                    source_kind="single_numeric",
                    reference_date=ref,
                )
                self._mark_auto_completed_year(fallback, fallback_auto_year, ref)
                if fallback.is_calendar_valid:
                    return fallback
                return self._invalid_numeric_result(parsed.status_text, parsed.status_code, ref)

            elif len(txt) in {5, 7}:
                result.status_text = STATUS_AMBIGUOUS_NUMERIC_DATE
                result.status_code = "ambiguous_numeric_date"
                return result

            else:
                result.status_text = STATUS_INVALID_LENGTH
                result.status_code = "invalid_length"
                return result

        except Exception:
            result.status_text = STATUS_UNCLEAR_DATE
            result.status_code = "unclear_date"
            return result

    def _invalid_numeric_result(
        self,
        status_text: str,
        status_code: str,
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        result = self._blank_result(source_kind="single_numeric", reference_date=reference_date)
        result.status_text = status_text or STATUS_UNCLEAR_DATE
        result.status_code = status_code or "unclear_date"
        return result

    def _normalize_date_separators(self, txt: str) -> str:
        text = str(txt).strip()
        text = re.sub(r"([./-])\1+", r"\1", text)
        return text.replace(".", "/").replace("-", "/")

    def _recover_trailing_text_date(
        self,
        txt: str,
        pattern: DateFormatPattern,
        reference_date: date,
    ) -> Optional[DateParseResult]:
        match = re.match(r"^\s*(\d{1,4}[./-]+\d{1,2}(?:[./-]+\d{2,4})?)([^\d./-].*)$", txt)
        if not match:
            return None

        candidate = self._normalize_date_separators(match.group(1))
        parsed = self._parse_separated_date_string(candidate, pattern, reference_date=reference_date)
        if parsed.is_calendar_valid:
            parsed.status_text = STATUS_TRAILING_TEXT_IGNORED
            parsed.status_code = "trailing_text_ignored"
            parsed.severity = "warning"
            return parsed
        return None

    # ----------------------------------------------------
    # SEPARATED DATE
    # ----------------------------------------------------

    # הפונקציה מיישמת parsing פנימי של תאריך עם מפרידים ועם דפוס גיליון.
    def _parse_separated_date_string(
        self,
        txt: str,
        pattern: DateFormatPattern,
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        ref = self._get_reference_date(reference_date)
        result = self._blank_result(source_kind="single_separated", reference_date=ref)

        trailing_recovered = self._recover_trailing_text_date(txt, pattern, ref)
        if trailing_recovered is not None:
            return trailing_recovered

        parts = txt.split("/")

        year_was_defaulted = False
        if len(parts) == 2 and all(p.isdigit() for p in parts):
            parts = [parts[0], parts[1], str(ref.year)]
            year_was_defaulted = True

        if len(parts) != 3 or not all(p.isdigit() for p in parts):
            result.status_text = STATUS_INVALID_FORMAT
            result.status_code = "invalid_format"
            return result

        try:
            if pattern == DateFormatPattern.MMDD:
                mo = int(parts[0])
                dy = int(parts[1])
            else:
                dy = int(parts[0])
                mo = int(parts[1])

            raw_year = int(parts[2])
            yr = raw_year
            auto_year: Optional[int] = None
            if yr < 100:
                auto_year = yr
                yr = self._expand_two_digit_year(yr, reference_date=ref)

            parsed = self._validate_date(yr, mo, dy, source_kind="single_separated", reference_date=ref)
            if auto_year is not None:
                self._mark_auto_completed_year(parsed, auto_year, ref)
            if year_was_defaulted:
                parsed.year_was_defaulted = True
                parsed.status_text = STATUS_MISSING_YEAR_DEFAULTED
                parsed.status_code = "missing_year_defaulted"
            return parsed

        except Exception:
            result.status_text = STATUS_UNCLEAR_DATE
            result.status_code = "unclear_date"
            return result

    # ----------------------------------------------------
    # MIXED MONTH-NUMERIC
    # ----------------------------------------------------

    # הפונקציה מטפלת בתאריכים הכוללים שם חודש וטקסט מספרי מעורב.
    def _parse_mixed_month_numeric(
        self,
        txt: str,
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        ref = self._get_reference_date(reference_date)
        result = self._blank_result(source_kind="single_month_name", reference_date=ref)

        month_num = self._extract_month_number(txt)
        if month_num == 0:
            result.status_text = STATUS_UNPARSEABLE
            result.status_code = "unparseable"
            return result

        tokens = re.split(r"[^\d]+", txt)
        nums = [int(t) for t in tokens if t.isdigit()]

        if len(nums) < 2:
            result.status_text = STATUS_MISSING_DAY
            result.status_code = "missing_day"
            result.missing_components = ["day"]
            return result

        yr = 0
        dy = 0
        auto_year: Optional[int] = None

        for n in nums:
            if 1000 <= n <= 9999:
                yr = n
                break

        remaining = [n for n in nums if n != yr]
        if not remaining:
            result.status_text = STATUS_UNPARSEABLE
            result.status_code = "unparseable"
            return result

        big = [n for n in remaining if n > 12]
        if big:
            dy = big[0]
        else:
            dy = remaining[0]

        if yr == 0:
            two_digits = [n for n in remaining if 0 <= n <= 99 and n != dy]
            if two_digits:
                auto_year = two_digits[0]
                yr = self._expand_two_digit_year(auto_year, reference_date=ref)

        if yr == 0 or dy == 0:
            result.status_text = STATUS_UNPARSEABLE
            result.status_code = "unparseable"
            return result

        parsed = self._validate_date(yr, month_num, dy, source_kind="single_month_name", reference_date=ref)
        if auto_year is not None:
            self._mark_auto_completed_year(parsed, auto_year, ref)
        return parsed

    # הפונקציה מזהה האם הטקסט כולל שם חודש במקום מספר חודש.
    def _contains_month_name(self, txt: str) -> bool:
        return self._extract_month_number(txt) != 0

    # הפונקציה ממירה שם חודש או וריאציה טקסטואלית למספר חודש.
    def _extract_month_number(self, txt: str) -> int:
        t = txt.lower()

        english_months = {
            "january": 1, "jan": 1,
            "february": 2, "feb": 2,
            "march": 3, "mar": 3,
            "april": 4, "apr": 4,
            "may": 5,
            "june": 6, "jun": 6,
            "july": 7, "jul": 7,
            "august": 8, "aug": 8,
            "september": 9, "sep": 9,
            "october": 10, "oct": 10,
            "november": 11, "nov": 11,
            "december": 12, "dec": 12,
        }
        for key, val in english_months.items():
            if key in t:
                return val

        hebrew_months = {
            "ינואר": 1,
            "פברואר": 2,
            "מרץ": 3,
            "מרס": 3,
            "אפריל": 4,
            "מאי": 5,
            "יוני": 6,
            "יולי": 7,
            "אוגוסט": 8,
            "ספטמבר": 9,
            "אוקטובר": 10,
            "נובמבר": 11,
            "דצמבר": 12,
        }
        for key, val in hebrew_months.items():
            if key in t:
                return val

        return 0

    # ----------------------------------------------------
    # VALIDATE DATE
    # ----------------------------------------------------

    # הפונקציה מאמתת רכיבי שנה/חודש/יום ומחזירה תוצאה תקנית או שגיאה.
    def _validate_date(
        self,
        yr,
        mo,
        dy,
        *,
        source_kind: str = "",
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        result = self._blank_result(source_kind=source_kind, reference_date=reference_date)

        try:
            yr = int(yr)
            mo = int(mo)
            dy = int(dy)
        except (TypeError, ValueError):
            result.status_text = STATUS_UNPARSEABLE
            result.status_code = "unparseable"
            return result

        result.year = yr
        result.month = mo
        result.day = dy

        if dy < 1 or dy > 31:
            result.status_text = STATUS_INVALID_DAY
            result.status_code = "invalid_day"
            result.invalid_components = ["day"]
            return result

        if mo < 1 or mo > 12:
            result.status_text = STATUS_INVALID_MONTH
            result.status_code = "invalid_month"
            result.invalid_components = ["month"]
            return result

        if yr < 1:
            result.status_text = STATUS_INVALID_YEAR
            result.status_code = "invalid_year"
            result.invalid_components = ["year"]
            return result

        # Reject years that are clearly outside any plausible business domain.
        # Python's datetime supports years 1–9999, but years beyond the
        # reference year + 1 are impossible for birth/entry dates.
        # We use 9999 as the hard ceiling to avoid datetime() raising ValueError
        # for astronomically large values (e.g. 1234567).
        ref = self._get_reference_date(reference_date)
        if yr > 9999 or yr > ref.year + 1:
            result.status_text = STATUS_IMPOSSIBLE_YEAR
            result.status_code = "impossible_year"
            result.invalid_components = ["year"]
            return result

        try:
            _ = datetime(yr, mo, dy)
        except Exception:
            result.status_text = STATUS_DATE_NOT_EXISTS
            result.status_code = "date_not_exists"
            result.invalid_components = ["day"]
            return result

        result.is_valid = True
        result.is_calendar_valid = True
        result.is_business_valid = True
        result.severity = "ok"
        result.status_code = "ok"
        return result

    # ----------------------------------------------------
    # BUSINESS RULES
    # ----------------------------------------------------

    # הפונקציה מריצה חוקים עסקיים על תאריך לאחר parsing, כולל גיל ותאריך עתידי.
    def validate_business_rules(
        self,
        result: DateParseResult,
        field_type: DateFieldType,
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        ref = self._get_reference_date(reference_date)
        result.reference_year = ref.year
        status_before_business = result.status_text

        if result.is_valid and not result.is_calendar_valid:
            result.is_calendar_valid = True
            result.is_business_valid = True

        if field_type == DateFieldType.ENTRY_DATE and result.status_text == STATUS_EMPTY_CELL:
            result.status_text = ""
            result.status_code = "empty_optional_entry"
            result.is_valid = False
            result.is_calendar_valid = False
            result.is_business_valid = False
            result.severity = "ok"
            return result

        if not result.is_calendar_valid:
            result.is_valid = False
            result.is_business_valid = False
            return result

        if result.year is None or result.month is None or result.day is None:
            result.is_valid = False
            result.is_business_valid = False
            return result

        if result.status_code == "missing_year_defaulted":
            result.is_valid = True
            result.is_business_valid = True
            result.severity = "warning"
            return result

        if result.year < 1906:
            result.is_valid = False
            result.is_business_valid = False
            result.severity = "error"
            self._set_status_preserving_existing(result, STATUS_BEFORE_1906)
            result.status_code = "year_before_1906"
            return result

        # Hard sanity check: reject years that are clearly impossible in any
        # business domain (far future, astronomical values, etc.).
        # This catches values like 5280, 2229, 9999 that pass calendar
        # validation but are nonsensical as birth or entry years.
        if result.year > ref.year + 1:
            result.is_valid = False
            result.is_business_valid = False
            result.severity = "error"
            self._set_status_preserving_existing(result, STATUS_IMPOSSIBLE_YEAR)
            result.status_code = "impossible_year"
            return result

        try:
            date_val = date(result.year, result.month, result.day)
        except Exception:
            return result

        if field_type == DateFieldType.ENTRY_DATE and result.year >= ref.year:
            result.is_valid = False
            result.is_business_valid = False
            result.severity = "error"
            self._set_status_preserving_existing(result, STATUS_LATE_ENTRY)
            result.status_code = "late_entry"
            return result

        if date_val > ref:
            result.is_valid = False
            result.is_business_valid = False
            result.severity = "error"
            if field_type == DateFieldType.BIRTH_DATE:
                self._set_status_preserving_existing(result, STATUS_FUTURE_BIRTH)
                result.status_code = "future_birth"
            else:
                self._set_status_preserving_existing(result, STATUS_FUTURE_ENTRY)
                result.status_code = "future_entry"
            return result

        if field_type == DateFieldType.BIRTH_DATE:
            age = self._calculate_age(date_val, ref)
            if age > 100:
                result.status_text = f"גיל מעל 100 ({age} שנים)"
                if status_before_business and status_before_business not in result.status_text:
                    result.status_text = f"{status_before_business} | {result.status_text}"
                result.status_code = "age_over_100"
                result.severity = "warning"
                result.is_valid = True
                result.is_business_valid = True
                return result

        result.is_valid = True
        result.is_business_valid = True
        if not result.status_text:
            result.severity = "ok"
            result.status_code = "ok"
        return result

    # ----------------------------------------------------
    # ENTRY BEFORE BIRTH
    # ----------------------------------------------------

    # הפונקציה בודקת שתאריך כניסה אינו קודם לתאריך הלידה באותה שורה.
    def validate_entry_before_birth(
        self,
        birth: DateParseResult,
        entry: DateParseResult
    ) -> bool:
        if not birth.is_calendar_valid or not entry.is_calendar_valid:
            return True

        if birth.year is None or birth.month is None or birth.day is None:
            return True

        if entry.year is None or entry.month is None or entry.day is None:
            return True

        try:
            birth_date = datetime(birth.year, birth.month, birth.day)
            entry_date = datetime(entry.year, entry.month, entry.day)
        except Exception:
            return True

        if entry_date < birth_date:
            logger.error(
                "Logical error: Entry date before birth date "
                "(Birth: %s, Entry: %s)",
                birth_date.date(),
                entry_date.date(),
            )
            return False

        return True

    # ----------------------------------------------------
    # HELPERS
    # ----------------------------------------------------

    # הפונקציה היא helper פנימי להרחבת שנה דו־ספרתית עם טיפול בערכים ריקים.
    def _expand_two_digit_year(self, yr, reference_date: Optional[date] = None):
        ref = self._get_reference_date(reference_date)
        current_two = ref.year % 100
        yr = int(yr)

        if yr <= current_two:
            return (ref.year // 100) * 100 + yr
        return ((ref.year // 100) - 1) * 100 + yr

    # הפונקציה בודקת האם לפחות אחד מרכיבי התאריך המפוצל קיים.
    def _has_any_split_component(self, y, m, d) -> bool:
        return not (self._is_empty(y) and self._is_empty(m) and self._is_empty(d))

    # הפונקציה משמרת תאימות לשם הישן ובודקת האם תאריך מפוצל מלא.
    def _has_split_date(self, y, m, d):
        return self._has_any_split_component(y, m, d)

    # הפונקציה מזהה ערך ריק בצורה אחידה עבור רכיבי תאריך.
    def _is_empty(self, value) -> bool:
        return value is None or str(value).strip() == ""

    def _is_split_year_zero(self, value) -> bool:
        if self._is_empty(value):
            return False
        try:
            return int(float(str(value).strip())) == 0
        except Exception:
            return False

    # הפונקציה ממירה רכיב תאריך מפוצל למספר ומסמנת אם ההמרה הצליחה.
    def _coerce_split_component(
        self,
        value,
        *,
        zero_is_missing: bool = False,
    ) -> tuple[Optional[int], bool, bool]:
        if self._is_empty(value):
            return None, False, True
        try:
            coerced = int(float(str(value).strip()))
            if zero_is_missing and coerced == 0:
                return None, False, True
            return coerced, True, False
        except Exception:
            return None, False, False

    # הפונקציה מחשבת גיל מדויק לפי תאריך לידה ותאריך ייחוס.
    def _calculate_age(self, birth: date, today: date) -> int:
        age = today.year - birth.year
        try:
            birthday_this_year = date(today.year, birth.month, birth.day)
        except Exception:
            return age

        if birthday_this_year > today:
            age -= 1
        return age

    def _get_reference_date(self, reference_date: Optional[date]) -> date:
        if reference_date is not None:
            return reference_date
        return self.reference_date

    def _set_status_preserving_existing(self, result: DateParseResult, status_text: str) -> None:
        existing = (result.status_text or "").strip()
        if existing and status_text and status_text not in existing:
            result.status_text = f"{existing} | {status_text}"
        else:
            result.status_text = status_text

    def _mark_auto_completed_year(self, result: DateParseResult, original_year: int, ref: date) -> None:
        result.year_was_auto_completed = True
        result.original_year_value = int(original_year)
        result.original_year_digits = 1 if int(original_year) < 10 else 2
        result.reference_year = ref.year

    def _missing_components_status(self, components: list[str]) -> str:
        if components == ["year"]:
            return STATUS_MISSING_YEAR
        if components == ["month"]:
            return STATUS_MISSING_MONTH
        if components == ["day"]:
            return STATUS_MISSING_DAY
        if components == ["month", "day"]:
            return STATUS_MISSING_MONTH_DAY
        translated = {
            "year": STATUS_MISSING_YEAR,
            "month": STATUS_MISSING_MONTH,
            "day": STATUS_MISSING_DAY,
        }
        return " | ".join(translated[c] for c in components if c in translated)

    def _looks_like_excel_serial(self, value: int) -> bool:
        # Avoid treating common explicit years as serial dates unless future
        # extraction metadata later confirms the cell was date-formatted.
        if 1900 <= value <= 2100:
            return False
        return 1 <= value <= 2958465

    def _parse_excel_serial_date(
        self,
        raw_value: int | float,
        reference_date: Optional[date] = None,
    ) -> DateParseResult:
        ref = self._get_reference_date(reference_date)
        result = self._blank_result(source_kind="single_numeric", reference_date=ref)

        try:
            from openpyxl.utils.datetime import from_excel

            dt = from_excel(raw_value)
            if isinstance(dt, datetime):
                dt = dt.date()
        except Exception:
            result.status_text = STATUS_EXCEL_SERIAL_NOT_RECOGNIZED
            result.status_code = "unrecognized_numeric_date"
            result.severity = "error"
            return result

        if not isinstance(dt, date):
            result.status_text = STATUS_EXCEL_SERIAL_NOT_RECOGNIZED
            result.status_code = "unrecognized_numeric_date"
            result.severity = "error"
            return result

        if dt.year < 1900 or dt.year > ref.year + 1:
            result.status_text = STATUS_EXCEL_SERIAL_NOT_RECOGNIZED
            result.status_code = "unrecognized_numeric_date"
            result.severity = "error"
            return result

        result.year = dt.year
        result.month = dt.month
        result.day = dt.day
        result.is_valid = True
        result.is_calendar_valid = True
        result.is_business_valid = True
        result.severity = "ok"
        result.status_code = "excel_serial_parsed"
        result.status_text = STATUS_EXCEL_SERIAL_PARSED
        return result

    def _looks_like_full_date_value(self, value) -> bool:
        if isinstance(value, (datetime, date)):
            return True
        if value is None:
            return False
        text = str(value).strip()
        return "/" in text or "." in text or bool(re.match(r"^\d{4}-\d{2}-\d{2}", text))

    def _split_full_date_status(self, column: str) -> str:
        if column == "day":
            return STATUS_SPLIT_FULL_DATE_FROM_DAY
        if column == "month":
            return STATUS_SPLIT_FULL_DATE_FROM_MONTH
        return STATUS_SPLIT_FULL_DATE_FROM_YEAR
