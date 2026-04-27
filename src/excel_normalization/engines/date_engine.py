"""DateEngine — date field parsing and validation rules.

Purpose:
    Parses date values from Excel cells (split year/month/day columns or a
    single combined cell) and validates them against business rules.  Returns
    a DateParseResult with the parsed components, a validity flag, and a
    Hebrew status message.

Implemented rules:

    1. Input routing: split columns vs. single value
       If year, month, and day values are all present and numeric, the split-
       column path is used.  Otherwise the single-value path is used.
       Example (split):
           year=1985, month=3, day=14  → parsed directly
       Example (single):
           "14/03/1985"  → parsed via separator logic

    2. Two-digit year expansion (split columns and separated strings)
       Years below 100 are expanded to a 4-digit year using a sliding window
       relative to the current year:
       - If the 2-digit year ≤ last two digits of current year → current century.
       - Otherwise → previous century.
       Example (assuming current year is 2026):
           yr=26  → 2026   (≤ 26, current century)
           yr=27  → 1927   (> 26, previous century)
           yr=85  → 1985   (> 26, previous century)
           yr=0   → 2000

    3. Numeric string parsing (no separator, all digits)
       8 digits: interpreted as DDMMYYYY.
           "14031985"  → day=14, month=3, year=1985
       6 digits: interpreted as DDMMYY (year expanded via rule 2).
           "140385"    → day=14, month=3, year=1985
       4 digits: if the value is a valid year (1900–2100) → year only,
           status="חסר חודש ויום", is_valid=False.
           "1985"      → year=1985, month=0, day=0, is_valid=False
           Otherwise interpreted as DMYY (1 digit day, 1 digit month, 2 digit year).
           "1385"      → day=1, month=3, year=1985

    4. Separated string parsing (contains "/" or ".")
       Dots are converted to slashes first.  Then split on "/".
       Two-part dates (DD/MM): current year is assumed.
           "14/03"     → day=14, month=3, year=<current year>
       Three-part dates (DD/MM/YYYY or MM/DD/YYYY):
           Default format (DDMM): first part=day, second=month.
               "14/03/1985"  → day=14, month=3, year=1985
           MMDD format (when pattern=DateFormatPattern.MMDD):
               "03/14/1985"  → month=3, day=14, year=1985
           Two-digit year in third part is expanded via rule 2.
               "14/03/85"    → day=14, month=3, year=1985

    5. ISO-like string parsing
       Strings matching YYYY-MM-DD (optionally with time suffix) are parsed
       directly.
           "1985-03-14"           → year=1985, month=3, day=14
           "1985-03-14T00:00:00"  → year=1985, month=3, day=14

    6. Excel date/datetime objects
       Python date or datetime objects passed directly are used as-is.
           date(1985, 3, 14)  → year=1985, month=3, day=14

    7. Excel serial date integers
       Integer values in the range 1–2958465 are treated as Excel serial
       dates and converted using openpyxl's from_excel() utility.
           36526  → 2000-01-01

    8. Month name parsing (English and Hebrew)
       Strings containing a month name (English or Hebrew) are parsed by
       extracting the month number and finding the remaining numeric tokens
       for day and year.  A number > 12 in the remaining tokens is preferred
       as the day value.
       Supported English names: january/jan, february/feb, march/mar,
           april/apr, may, june/jun, july/jul, august/aug, september/sep,
           october/oct, november/nov, december/dec.
       Supported Hebrew names: ינואר, פברואר, מרץ, מרס, אפריל, מאי, יוני,
           יולי, אוגוסט, ספטמבר, אוקטובר, נובמבר, דצמבר.
       Example:
           "14 March 1985"   → day=14, month=3, year=1985
           "14 מרץ 1985"     → day=14, month=3, year=1985

    9. Date component validation
       After parsing, components are validated:
       - Day must be 1–31; otherwise status="יום לא תקין", is_valid=False.
       - Month must be 1–12; otherwise status="חודש לא תקין", is_valid=False.
       - Year must be ≥ 1; otherwise status="שנה לא תקינה", is_valid=False.
       - The combination must form a real calendar date (e.g. Feb 30 is
         rejected); otherwise status="תאריך לא קיים", is_valid=False.
       Note: even when is_valid=False, the parsed year/month/day components
       are stored in the result so callers can display them.

    10. Business rules — birth date (DateFieldType.BIRTH_DATE)
        Applied after component validation:
        - Year < 1900 → status="שנה לפני 1900", is_valid=False.
        - Date is in the future → status="תאריך לידה עתידי", is_valid=False.
        - Age > 100 years → status="גיל מעל 100 (N שנים)", is_valid stays True
          (observed behavior: age warning does not invalidate the date).
        Age is calculated exactly: birthday not yet reached this year subtracts 1.

    11. Business rules — entry date (DateFieldType.ENTRY_DATE)
        Applied after component validation:
        - Empty entry date → status="", is_valid=False (not an error).
        - Year < 1900 → status="שנה לפני 1900", is_valid=False.
        - Year ≥ current year → status="תאריך כניסה מאוחר מהתאריך שנקבע",
          is_valid=False.  This is stricter than the birth date future check:
          any date in the current year is rejected, even if it has already
          passed within the year.
        - Date is in the future (and year < current year) →
          status="תאריך כניסה עתידי", is_valid=False.

    12. Entry date before birth date (cross-field validation)
        validate_entry_before_birth() checks whether the entry date precedes
        the birth date.  Returns False when entry < birth (both valid).
        This method is called by the processor layer, not internally.
        The actual status cell update (Hebrew warning appended to entry status)
        is done by the orchestrator, not by this engine.

Status messages returned (Hebrew):
    "תא ריק"                              — empty cell
    "תוכן לא ניתן לפריקה"                 — value cannot be parsed
    "פורמט תאריך לא מזוהה"                — unrecognized format
    "פורמט תאריך לא תקין"                 — invalid format
    "אורך תאריך לא תקין"                  — numeric string wrong length
    "תאריך לא ברור"                       — ambiguous date
    "אין מפריד בתאריך"                    — no separator found
    "חסר חודש ויום"                       — only year present
    "חסר יום"                             — month name found but no day
    "יום לא תקין"                         — day out of range
    "חודש לא תקין"                        — month out of range
    "שנה לא תקינה"                        — year < 1
    "תאריך לא קיים"                       — calendar date does not exist
    "שנה לפני 1900"                       — year before 1900
    "תאריך לידה עתידי"                    — birth date in the future
    "תאריך כניסה עתידי"                   — entry date in the future
    "תאריך כניסה מאוחר מהתאריך שנקבע"    — entry year ≥ current year
    "גיל מעל 100 (N שנים)"               — age exceeds 100 years
    "" (empty)                            — valid date or empty entry date

Important notes:
    - DateEngine does not read or write Excel files.
    - It does not call any web or I/O layer.
    - The engine operates on values already extracted from Excel cells.
    - parse_date() is the main entry point; it routes to split or single
      parsing and then applies business rules.
    - DateFormatPattern (DDMM vs. MMDD) is determined by the processor layer
      from the column header context, not by this engine.

Known limitations:
    - The MMDD format detection is not automatic; it must be passed in by the
      caller.  If the wrong pattern is passed, day and month will be swapped.
    - Two-part dates (DD/MM without year) assume the current year, which may
      produce incorrect results for historical data.
    - The age > 100 check sets a status message but does not set is_valid=False,
      so the value is still written to the corrected column.
    - The entry date cutoff (year ≥ current year) is calendar-year based, not
      day-based: a January 2026 entry date is rejected when checked in April 2026.
"""

from datetime import date, datetime
import logging
import re

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
        year_was_auto_completed = yr < 100

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

        if result.year < 1900:

            result.is_valid = False
            result.status_text = "שנה לפני 1900"
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

        try:
            if y in (None, "") or m in (None, "") or d in (None, ""):
                return False
            # Accept numeric types directly (int/float from Excel cells)
            if isinstance(y, (int, float)) and isinstance(m, (int, float)) and isinstance(d, (int, float)):
                return True
            return (
                str(y).strip().replace(".", "").replace("-", "").isdigit()
                and str(m).strip().replace(".", "").replace("-", "").isdigit()
                and str(d).strip().replace(".", "").replace("-", "").isdigit()
            )
        except Exception:
            return False

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