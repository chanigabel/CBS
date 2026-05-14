"""Israeli ID and passport normalization rules.

The engine cleans ID values, applies checksum validation, preserves passport
values when appropriate, and routes passport-like nonnumeric ID values to
passport when the business rules require it.
"""

import logging
from typing import Any, Tuple

from ..data_types import IdentifierResult

logger = logging.getLogger(__name__)


# המנוע אחראי לניקוי, סיווג ואימות תעודות זהות ודרכונים.
class IdentifierEngine:
    engine_key = "identifier"
    display_name = "Identifier Engine"
    version = "1.0.0"
    description = "Validates ID values and normalizes passport fields."
    supported_fields = ["id_number", "passport"]

    """Pure business logic for ID and passport validation.

    This class contains no Excel dependencies and operates on plain Python types.
    All methods are deterministic and replicate VBA behavior exactly.
    """

    # Dash Unicode characters accepted in IDs
    DASH_CHARS = {
        45,  # hyphen
        8209,  # non-breaking hyphen
        8210,  # figure dash
        8211,  # en-dash
        8212,  # em-dash
        8213,  # horizontal bar
        8722,  # minus sign
    }

    # הפונקציה מנרמלת זוג שדות מזהה/דרכון ומחזירה תוצאה אחת לשורת ה־Dataset.
    def normalize_identifiers(self, id_value: Any, passport_value: Any) -> IdentifierResult:
        """Process ID and passport values together (VBA NormalizeIdentifiers parity)."""
        id_str = self._safe_to_string(id_value).strip()
        passport_str = self._safe_to_string(passport_value).strip()

        # Clean passport first
        cleaned_passport = self.clean_passport(passport_str)

        # Treat "9999" as no ID (sentinel check on original string before cleanup)
        if id_str == "9999":
            logger.info("Special ID value '9999' treated as no ID (Status: חסר מזהים)")
            id_str = ""

        # No ID provided
        if not id_str:
            if cleaned_passport:
                status = "דרכון הוזן"
            else:
                status = "חסר מזהים"
            return IdentifierResult(corrected_id="", corrected_passport=cleaned_passport, status_text=status)

        # Remove hyphens from the ID field before validation.
        # Hyphens are the only separator stripped here; letters, spaces, and
        # other characters are left intact so _process_id_value can decide
        # whether to route the value to the passport field.
        id_str_digits = self.clean_id_number(id_str)

        # If nothing remains after hyphen removal (e.g. "---"), treat as missing ID
        if not id_str_digits:
            if cleaned_passport:
                status = "ת.ז. לא תקינה + דרכון הוזן"
            else:
                status = "ת.ז. לא תקינה"
            return IdentifierResult(corrected_id="", corrected_passport=cleaned_passport, status_text=status)

        # Process the hyphen-stripped form through the existing Israeli ID logic.
        # Letters and special characters remain so _process_id_value can route
        # passport-like values to passport when allowed.
        (
            cleaned_digits,
            moved_to_passport,
            updated_passport,
            checksum_valid,
        ) = self._process_id_value(id_str_digits, cleaned_passport)

        cleaned_passport = updated_passport

        # If ID was moved to passport because it contains passport-like
        # nonnumeric content.
        if moved_to_passport:
            _digits, _should_move, reason = self.classify_id_value(id_str_digits)
            # ID column is empty in output, passport gets value
            if reason in {"too_short", "too_long"}:
                status = "ת.ז. לא תקינה + הועברה לדרכון"
            else:
                status = "ת.ז. הועברה לדרכון"
            return IdentifierResult(corrected_id="", corrected_passport=cleaned_passport, status_text=status)

        # At this point, we did not move the ID into passport. cleaned_digits may
        # be empty (e.g., all zeros) or 9-digit string used only for validation.

        if not cleaned_digits:
            # No usable digits but we decided not to move to passport.
            if cleaned_passport:
                status = "ת.ז. לא תקינה + דרכון הוזן"
            else:
                status = "ת.ז. לא תקינה"
            _digits, _should_move, reason = self.classify_id_value(id_str_digits)
            if reason in {"too_short", "too_long"}:
                return IdentifierResult(corrected_id="", corrected_passport=cleaned_passport, status_text=status)
            return IdentifierResult(corrected_id=id_str, corrected_passport=cleaned_passport, status_text=status)

        # Valid or invalid checksum with a 9-digit cleaned ID
        if checksum_valid:
            if cleaned_passport:
                status = "ת.ז. תקינה + דרכון הוזן"
            else:
                status = "ת.ז. תקינה"
            # Return the hyphen-stripped form as corrected_id so the output
            # column shows a clean value without separator characters.
            return IdentifierResult(corrected_id=id_str_digits, corrected_passport=cleaned_passport, status_text=status)

        # Invalid checksum
        if cleaned_passport:
            status = "ת.ז. לא תקינה + דרכון הוזן"
        else:
            status = "ת.ז. לא תקינה"
        # For invalid IDs that were not moved, VBA commonly keeps the normalized/padded
        # numeric value in the corrected ID column (while valid IDs keep the original).
        return IdentifierResult(corrected_id=cleaned_digits, corrected_passport=cleaned_passport, status_text=status)

    # ------------------------------------------------------------------
    # Internal helpers mirroring VBA ProcessIDValue
    # ------------------------------------------------------------------

    # הפונקציה ממירה ערך קלט לטקסט בטוח לפני ניקוי מזהים.
    def _safe_to_string(self, v: Any) -> str:
        try:
            return "" if v is None else str(v)
        except Exception:
            return ""

    # הפונקציה משאירה רק ספרות עבור בדיקות תעודת זהות.
    def _clean_digits_only(self, txt: str) -> str:
        return "".join(ch for ch in txt if ch.isdigit())

    # הפונקציה מנקה מספר זהות ומכינה אותו לסיווג ולאימות checksum.
    def clean_id_number(self, id_str: str) -> str:
        """Remove hyphen characters from an ID field value before validation.

        Israeli ID numbers are purely numeric, but they are sometimes written
        with hyphens as separators (e.g. "039-337-423").  Hyphens are the only
        separator that should be silently removed; all other non-digit characters
        (letters, spaces, dots, etc.) are left in place so that the downstream
        character scan in _process_id_value can decide whether to move the value
        to the passport field.

        This is intentionally narrow: only hyphens (all variants in DASH_CHARS)
        are stripped.  A value like "A218988699" is NOT digit-cleaned here —
        the letter causes _process_id_value to route it to passport, which is
        the correct behaviour for alphanumeric values.

        Args:
            id_str: Raw ID string (already stripped of leading/trailing whitespace).

        Returns:
            id_str with all hyphen-like characters removed.
        """
        return "".join(ch for ch in id_str if ord(ch) not in self.DASH_CHARS)

    # הפונקציה מעבדת ערך מזהה יחיד ומחליטה אם הוא תעודת זהות, דרכון או שגיאה.
    def _process_id_value(
        self,
        id_str: str,
        cleaned_passport: str,
    ) -> Tuple[str, bool, str, bool]:
        """Replicates VBA ProcessIDValue for a single ID.

        Returns:
            (cleaned_digits, moved_to_passport, updated_passport, checksum_valid)
        """
        # Scan characters – any non-digit, non-dash => move to passport
        for ch in id_str:
            if not ch.isdigit() and ord(ch) not in self.DASH_CHARS:
                # Move to passport only if passport currently empty; otherwise keep it
                if not cleaned_passport:
                    cleaned_passport = self.clean_passport(id_str)
                logger.warning("ID moved to passport: %r (Reason: Contains non-digit characters)", id_str)
                # moved_to_passport = True, no cleaned_digits, checksum_valid=False
                return "", True, cleaned_passport, False

        # Extract digits
        digits = self._clean_digits_only(id_str)

        # All zeros => invalid, do not move
        if digits and all(ch == "0" for ch in digits):
            logger.info("All-zeros ID rejected: %r (Status: לא תקין)", digits)
            return "", False, cleaned_passport, False

        digit_count = len(digits)

        # Too few digits (<4) => invalid ID. Numeric-only values must not move
        # to passport just because of length.
        if digit_count < 4:
            logger.warning("Invalid ID length: %r (Reason: < 4 digits)", id_str)
            return digits, False, cleaned_passport, False

        # Too many digits (>9) => invalid ID. Numeric-only values must not move
        # to passport just because of length.
        if digit_count > 9:
            logger.warning("Invalid ID length: %r (Reason: > 9 digits)", id_str)
            return digits, False, cleaned_passport, False

        # 4–9 digits: pad to 9 and validate checksum
        padded = self.pad_id(digits)

        # Reject all-zero padded (already handled by digits check, but keep parity)
        if padded == "000000000":
            logger.info("All-zeros padded ID rejected: %r", padded)
            return "", False, cleaned_passport, False

        # Reject all-identical-digit padded sequences (e.g., "111111111")
        if len(set(padded)) == 1:
            logger.info("Identical-digit ID rejected: %r", padded)
            return "", False, cleaned_passport, False

        is_valid = self.validate_israeli_id(padded)
        if not is_valid:
            logger.warning("Invalid ID checksum: %r", padded)
        return padded, False, cleaned_passport, is_valid

    # ------------------------------------------------------------------
    # Public compatibility helpers used by unit tests / legacy callers
    # ------------------------------------------------------------------

    # הפונקציה מסווגת ערך מזהה לשימוש בדוחות ובבדיקות ממוקדות.
    def classify_id_value(self, id_value: Any) -> Tuple[str, bool, str]:
        """Classify an ID value for 'move-to-passport' decisions.

        Returns:
            (digits_only, should_move_to_passport, reason_code)

        reason_code values used by tests:
            - "" (no reason)
            - "invalid_format"
            - "too_short"
            - "too_long"
        """
        s = self._safe_to_string(id_value).strip()
        if s == "9999":
            s = ""
        if not s:
            return "", False, ""

        for ch in s:
            if not ch.isdigit() and ord(ch) not in self.DASH_CHARS:
                return "", True, "invalid_format"

        digits = self._clean_digits_only(s)
        if len(digits) < 4:
            return digits, False, "too_short"
        if len(digits) > 9:
            return digits, False, "too_long"
        return digits, False, ""

    # הפונקציה בודקת checksum של תעודת זהות ישראלית לאחר ניקוי ופדינג.
    def validate_israeli_id(self, id_digits: str) -> bool:
        """Validate Israeli ID checksum.

        Algorithm:
        - Multiply digits at odd positions (1st, 3rd, 5th, 7th, 9th) by 1
        - Multiply digits at even positions (2nd, 4th, 6th, 8th) by 2
        - If result > 9, subtract 9
        - Sum all results
        - Check if sum is divisible by 10

        Args:
            id_digits: 9-digit Israeli ID string

        Returns:
            True if checksum is valid, False otherwise

        VBA Equivalent: ValidateChecksum
        """
        if len(id_digits) != 9:
            return False

        checksum = 0
        for i, digit_char in enumerate(id_digits):
            digit = int(digit_char)

            # Position is 1-indexed: i=0 is position 1 (odd), i=1 is position 2 (even)
            if i % 2 == 0:  # Odd position (1, 3, 5, 7, 9)
                checksum += digit
            else:  # Even position (2, 4, 6, 8)
                val = digit * 2
                if val > 9:
                    val -= 9
                checksum += val

        return checksum % 10 == 0

    # הפונקציה משלימה תעודת זהות קצרה לאורכה התקני לפני בדיקה או יצוא.
    def pad_id(self, id_digits: str) -> str:
        """Pad ID to 9 digits with leading zeros.

        Args:
            id_digits: ID string with 4-9 digits

        Returns:
            9-digit ID string padded with leading zeros

        VBA Equivalent: Part of ProcessIDValue logic
        """
        return id_digits.zfill(9)

    # הפונקציה מנקה דרכון מערכי קלט לא אחידים ומכינה אותו לשדה היצוא.
    def clean_passport(self, passport: str) -> str:
        """Remove invalid characters from passport value.

        Keeps only:
        - Digits (0-9)
        - English letters (A-Z, a-z)
        - Hebrew letters (Unicode 1488-1514)
        - Dash characters (all variants)

        Args:
            passport: Raw passport value

        Returns:
            Cleaned passport value

        VBA Equivalent: CleanPassportValue
        """
        if not passport:
            return ""

        cleaned = []
        for char in passport:
            char_code = ord(char)

            # Check if digit
            if char.isdigit():
                cleaned.append(char)
            # Check if English letter
            elif char.isalpha() and char.isascii():
                cleaned.append(char)
            # Check if Hebrew letter (Unicode 1488-1514)
            elif 1488 <= char_code <= 1514:
                cleaned.append(char)
            # Check if dash character
            elif char_code in self.DASH_CHARS:
                cleaned.append(char)

        return "".join(cleaned)
