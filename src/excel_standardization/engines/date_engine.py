"""Date parsing and validation rules for birth and entry dates.

The engine parses split or combined date values, expands two-digit years,
validates calendar dates, and applies the birth-date and entry-date business
rules used by the active pipeline.
"""

from datetime import date, datetime
import logging
import re
from typing import Optional

from ..data_types import DateParseResult, DateFormatPattern, DateFieldType


logger = logging.getLogger(__name__)


class DateEngine:
    def _blank_result(self) -> DateParseResult:
        return DateParseResult(year=None, month=None, day=None, is_valid=False, status_text="")

    # ----------------------------------------------------
    # MAIN ENTRY
    # ----------------------------------------------------

    def parse_date(
        self,
        year_val,
        month_val,
        day_val,
        main_val,
        pattern: DateFormatPattern,
        field_type: DateFieldType,
    ) -> DateParseResult:

        if self._has_split_date(year_val, month_val, day_val):
            result = self.parse_from_split_columns(year_val, month_val, day_val)
        else:
            result = self.parse_from_main_value(main_val, pattern)

        return self.validate_business_rules(result, field_type)

    # ----------------------------------------------------
    # SPLIT COLUMNS
    # ----------------------------------------------------

    def parse_from_split_columns(self, year_val, month_val, day_val) -> DateParseResult:
        result = self._blank_result()

        yr, year_ok = self._coerce_split_component(year_val)
        mo, month_ok = self._coerce_split_component(month_val)
        dy, day_ok = self._coerce_split_component(day_val)

        result.year = yr
        result.month = mo
        result.day = dy

        if not (year_ok and month_ok and day_ok):
            result.status_text = "ערך תאריך לא תקין"
            return result

        try:
            yr = int(float(str(year_val).strip()))
            mo = int(float(str(month_val).strip()))
            dy = int(float(str(day_val).strip()))
        except Exception:
            result.status_text = "תוכן לא ניתן לפריקה"
            return result

        # Track whether the year was auto-completed from a shortened (< 100)
        # value.  The list-level majority correction in DateFieldProcessor
        # uses this flag to distinguish auto-completed years from explicitly
        # written 4-digit years.
        year_was_auto_completed = 0 <= yr < 100

        if year_was_auto_completed:
            yr = self._expand_two_digit_year(yr)

        result = self._validate_date(yr, mo, dy)
        result.year_was_auto_completed = year_was_auto_completed
        return result

    # ----------------------------------------------------
    # MAIN VALUE
    # ----------------------------------------------------

    def parse_from_main_value(
        self,
        raw_value,
        pattern: DateFormatPattern,
    ) -> DateParseResult:
        """Backward-compatible wrapper that now delegates to parse_date_value."""
        return self.parse_date_value(raw_value, pattern)

    def parse_date_value(self, raw_value, pattern: DateFormatPattern) -> DateParseResult:
        """Parse a date from a single cell value following VBA rules."""
        result = self._blank_result()

        if raw_value is None:
            result.status_text = "תא ריק"
            return result

        txt = str(raw_value).strip()
        if txt == "":
            result.status_text = "תא ריק"
            return result

        # Excel date/datetime
        if isinstance(raw_value, (datetime, date)):
            dt = raw_value if isinstance(raw_value, date) else raw_value.date()
            result.year = dt.year
            result.month = dt.month
            result.day = dt.day
            result.is_valid = True
            return result

        # Excel serial date number (integer, e.g. 36526 = 2000-01-01)
        # openpyxl with data_only=True sometimes returns these as integers
        if isinstance(raw_value, int) and 1 <= raw_value <= 2958465:
            try:
                from openpyxl.utils.datetime import from_excel
                dt = from_excel(raw_value)
                if isinstance(dt, datetime):
                    dt = dt.date()
                result.year = dt.year
                result.month = dt.month
                result.day = dt.day
                result.is_valid = True
                return result
            except Exception:
                pass  # Fall through to numeric string parsing

        # Contains month name (English or Hebrew)
        if self._contains_month_name(txt):
            return self._parse_mixed_month_numeric(txt)

        # All digits
        if txt.isdigit():
            return self._parse_numeric_date_string(txt)

        # ISO-like date string (common when merged date cells get stringified)
        # Example: "1997-09-04T00:00:00"
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", txt)
        if m:
            try:
                yr = int(m.group(1))
                mo = int(m.group(2))
                dy = int(m.group(3))
                return self._validate_date(yr, mo, dy)
            except Exception:
                # Fall through to standard parsing
                pass

        # Separated by "/" or "."
        if "/" in txt or "." in txt:
            txt2 = txt.replace(".", "/")
            return self._parse_separated_date_string(txt2, pattern)

        result.status_text = "פורמט תאריך לא מזוהה"
        return result

    # ------------------------------------------------------------------
    # Public compatibility wrappers (used by unit tests / legacy callers)
    # ------------------------------------------------------------------

    def expand_two_digit_year(self, year: int) -> int:
        return self._expand_two_digit_year(year)

    def parse_numeric_date_string(self, txt: str) -> DateParseResult:
        if txt is None:
            r = self._blank_result()
            r.status_text = "פורמט תאריך לא תקין"
            return r
        s = str(txt).strip()
        if not s.isdigit():
            r = self._blank_result()
            r.status_text = "פורמט תאריך לא תקין"
            return r
        return self._parse_numeric_date_string(s)

    def parse_separated_date_string(self, txt: str, pattern: DateFormatPattern) -> DateParseResult:
        if txt is None:
            r = self._blank_result()
            r.status_text = "אין מפריד בתאריך"
            return r
        s = str(txt).strip()
        if "/" not in s and "." not in s:
            r = self._blank_result()
            r.status_text = "אין מפריד בתאריך"
            return r
        s2 = s.replace(".", "/")
        return self._parse_separated_date_string(s2, pattern)

    def calculate_age(self, *args) -> int:
        """Compatibility wrapper.

        Supports:
        - calculate_age(birth: date, today: date)
        - calculate_age(birth_year: int, birth_month: int, birth_day: int)
        """
        if len(args) == 2 and isinstance(args[0], date) and isinstance(args[1], date):
            return self._calculate_age(args[0], args[1])
        if len(args) == 3:
            birth = date(int(args[0]), int(args[1]), int(args[2]))
            return self._calculate_age(birth, date.today())
        raise TypeError("calculate_age expects (birth, today) or (year, month, day)")

    # ----------------------------------------------------
    # NUMERIC DATE
    # ----------------------------------------------------

    def _parse_numeric_date_string(self, txt: str) -> DateParseResult:
        result = self._blank_result()

        try:

            if len(txt) == 8:

                dy = int(txt[0:2])
                mo = int(txt[2:4])
                yr = int(txt[4:8])

            elif len(txt) == 6:

                dy = int(txt[0:2])
                mo = int(txt[2:4])
                yr = self._expand_two_digit_year(int(txt[4:6]))

            elif len(txt) == 4:
                # Either a 4-digit year (YYYY) or DMYY (d m yy) VBA-style.
                yr_int = int(txt)
                if 1900 <= yr_int <= 2100:
                    result.year = yr_int
                    result.month = 0
                    result.day = 0
                    result.is_valid = False
                    result.status_text = "חסר חודש ויום"
                    return result

                dy = int(txt[0:1])
                mo = int(txt[1:2])
                yr = self._expand_two_digit_year(int(txt[2:4]))

            else:
                result.status_text = "אורך תאריך לא תקין"
                return result

            return self._validate_date(yr, mo, dy)

        except Exception:
            result.status_text = "תאריך לא ברור"
            return result

    # ----------------------------------------------------
    # SEPARATED DATE
    # ----------------------------------------------------

    def _parse_separated_date_string(
        self,
        txt: str,
        pattern: DateFormatPattern,
    ) -> DateParseResult:

        result = self._blank_result()

        parts = txt.split("/")

        # Two-part date: assume current year (common in forms)
        if len(parts) == 2 and all(p.isdigit() for p in parts):
            parts = [parts[0], parts[1], str(date.today().year)]

        if len(parts) != 3 or not all(p.isdigit() for p in parts):
            result.status_text = "פורמט תאריך לא תקין"
            return result

        try:

            if pattern == DateFormatPattern.MMDD:

                mo = int(parts[0])
                dy = int(parts[1])

            else:

                dy = int(parts[0])
                mo = int(parts[1])

            yr = int(parts[2])

            if yr < 100:
                yr = self._expand_two_digit_year(yr)

            return self._validate_date(yr, mo, dy)

        except Exception:

            result.status_text = "תאריך לא ברור"
            return result

    # ----------------------------------------------------
    # MIXED MONTH-NUMERIC (e.g., "12 January 2005", "ינואר 12 2005")
    # ----------------------------------------------------

    def _parse_mixed_month_numeric(self, txt: str) -> DateParseResult:
        result = self._blank_result()

        month_num = self._extract_month_number(txt)
        if month_num == 0:
            result.status_text = "תוכן לא ניתן לפריקה"
            return result

        tokens = re.split(r"[^\d]+", txt)
        nums = [int(t) for t in tokens if t.isdigit()]

        if len(nums) < 2:
            result.status_text = "חסר יום"
            return result

        yr = 0
        dy = 0

        # Prefer a 4-digit number as year
        for n in nums:
            if 1000 <= n <= 9999:
                yr = n
                break

        remaining = [n for n in nums if n != yr]
        if not remaining:
            result.status_text = "תוכן לא ניתן לפריקה"
            return result

        # Choose day from remaining numbers: >12 preferred
        big = [n for n in remaining if n > 12]
        if big:
            dy = big[0]
        else:
            dy = remaining[0]

        # If year still 0, look for 2-digit year candidate
        if yr == 0:
            two_digits = [n for n in remaining if 0 <= n <= 99 and n != dy]
            if two_digits:
                yr = self._expand_two_digit_year(two_digits[0])

        if yr == 0 or dy == 0:
            result.status_text = "תוכן לא ניתן לפריקה"
            return result

        return self._validate_date(yr, month_num, dy)

    def _contains_month_name(self, txt: str) -> bool:
        return self._extract_month_number(txt) != 0

    def _extract_month_number(self, txt: str) -> int:
        """Extract month number from text containing a month name."""
        t = txt.lower()

        english_months = {
            "january": 1,
            "jan": 1,
            "february": 2,
            "feb": 2,
            "march": 3,
            "mar": 3,
            "april": 4,
            "apr": 4,
            "may": 5,
            "june": 6,
            "jun": 6,
            "july": 7,
            "jul": 7,
            "august": 8,
            "aug": 8,
            "september": 9,
            "sep": 9,
            "october": 10,
            "oct": 10,
            "november": 11,
            "nov": 11,
            "december": 12,
            "dec": 12,
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

    def _validate_date(self, yr, mo, dy) -> DateParseResult:

        result = self._blank_result()

        # Coerce to int safely
        try:
            yr = int(yr)
            mo = int(mo)
            dy = int(dy)
        except (TypeError, ValueError):
            result.status_text = "תוכן לא ניתן לפריקה"
            return result

        # Always store the parsed components so callers can display them
        # even when the date is invalid.  is_valid stays False and
        # status_text carries the error description.
        result.year = yr
        result.month = mo
        result.day = dy

        if dy < 1 or dy > 31:
            result.status_text = "יום לא תקין"
            return result

        if mo < 1 or mo > 12:
            result.status_text = "חודש לא תקין"
            return result

        if yr < 1:
            result.status_text = "שנה לא תקינה"
            return result

        try:
            _ = datetime(yr, mo, dy)

        except ValueError:
            result.status_text = "תאריך לא קיים"
            return result
        except Exception:
            result.status_text = "תאריך לא קיים"
            return result

        result.is_valid = True

        return result

    # ----------------------------------------------------
    # BUSINESS RULES
    # ----------------------------------------------------

    def validate_business_rules(
        self,
        result: DateParseResult,
        field_type: DateFieldType
    ) -> DateParseResult:

        if field_type == DateFieldType.ENTRY_DATE and result.status_text == "תא ריק":
            # Empty entry date is considered valid and status must be cleared
            result.status_text = ""
            result.is_valid = False
            return result

        if not result.is_valid:
            return result

        today = date.today()

        # Minimum birth year is 1906 per institution-report requirements.
        # The status message says "שנה לפני 1906" for the validator-level check.
        # The DateEngine itself enforces 1906 here so that ALL paths (Excel/CLI,
        # web/JSON) produce the correct status without requiring the validator.
        if result.year < 1906:

            result.is_valid = False
            result.status_text = "שנה לפני 1906"
            return result

        try:

            date_val = date(result.year, result.month, result.day)

        except Exception:

            return result

        if field_type == DateFieldType.ENTRY_DATE:
            # Cutoff rule: entry_date.year must be <= current_year - 1.
            # This check takes priority over the generic "future date" check
            # because it is stricter: any date in the current year or later
            # is considered too late, even if it has already passed within
            # the current year (e.g. a January 2026 date checked in April 2026).
            if result.year >= today.year:
                result.is_valid = False
                result.status_text = "תאריך כניסה מאוחר מהתאריך שנקבע"
                return result

        if date_val > today:

            result.is_valid = False

            if field_type == DateFieldType.BIRTH_DATE:
                result.status_text = "תאריך לידה עתידי"
            else:
                result.status_text = "תאריך כניסה עתידי"

            return result

        if field_type == DateFieldType.BIRTH_DATE:

            # Exact age with birthday check
            age = self._calculate_age(date_val, today)

            if age > 100:
                result.status_text = f"גיל מעל 100 ({age} שנים)"

        return result

    # ----------------------------------------------------
    # ENTRY BEFORE BIRTH
    # ----------------------------------------------------

    def validate_entry_before_birth(
        self,
        birth: DateParseResult,
        entry: DateParseResult
    ) -> bool:

        if not birth.is_valid or not entry.is_valid:
            return True

        if not birth.year or not birth.month or not birth.day:
            return True

        if not entry.year or not entry.month or not entry.day:
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

    def _expand_two_digit_year(self, yr):

        current = date.today().year
        current_two = current % 100

        if yr <= current_two:
            return (current // 100) * 100 + yr
        else:
            return ((current // 100) - 1) * 100 + yr

    def _has_split_date(self, y, m, d):

        return (
            not self._is_empty(y)
            and not self._is_empty(m)
            and not self._is_empty(d)
        )

    def _is_empty(self, value) -> bool:
        return value is None or str(value).strip() == ""

    def _coerce_split_component(self, value) -> tuple[Optional[int], bool]:
        if self._is_empty(value):
            return None, False
        try:
            return int(float(str(value).strip())), True
        except Exception:
            return None, False

    def _calculate_age(self, birth: date, today: date) -> int:
        """Exact age calculation equivalent to VBA DateDiff('yyyy') with birthday check."""
        age = today.year - birth.year
        try:
            birthday_this_year = date(today.year, birth.month, birth.day)
        except Exception:
            return age

        if birthday_this_year > today:
            age -= 1
        return age
