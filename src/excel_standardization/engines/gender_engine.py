"""GenderEngine — gender field standardization rules.

Purpose:
    Normalizes gender values from any representation (Hebrew letters, Hebrew
    words, English words, numeric codes) to a standardized integer code:
    1 (male) or 2 (female).  Returns "" for unrecognized values.

Implemented rules:

    1. Input conversion
       Any input type (string, int, float, None) is converted to a string,
       stripped of leading/trailing whitespace, and lowercased before matching.
       Example:
           2       → "2"  → matched as female
           " נ "   → "נ"  → matched as female

    2. Female pattern matching (checked first)
       The normalized string is checked for any of the following substrings
       (case-insensitive):
           "2", "female", "נ", "אישה", "בת", "f", "נקבה", "girl", "woman"
       If any match is found → returns integer 2.
       Example:
           "נ"       → 2
           "נקבה"    → 2
           "female"  → 2
           "f"       → 2
           "2"       → 2

    3. Male pattern matching (checked only if female check fails)
       The normalized string is checked for any of the following substrings
       (case-insensitive):
           "1", "male", "ז", "זכר", "בן", "m", "man", "boy"
       If any match is found → returns integer 1.
       Note: female patterns are checked first to prevent "female" from
       matching the "m" inside it.
       Example:
           "ז"    → 1
           "זכר"  → 1
           "male" → 1
           "m"    → 1
           "1"    → 1

    4. None / empty input
       None → returns 1 (male, observed default behavior).
       Empty string after strip → returns 1 (male, observed default behavior).
       Note: the standardization pipeline short-circuits before calling this
       method for None/empty values, so this default is rarely reached in
       practice.

    5. Unrecognized values
       If the value does not match any female or male pattern → returns "".
       The empty string signals to the pipeline that the value could not be
       normalized and must not be copied as-is into the corrected field.
       Example:
           "8"    → ""
           "xyz"  → ""
           "?"    → ""

Important notes:
    - GenderEngine does not read or write Excel files.
    - It does not call any web or I/O layer.
    - Matching is substring-based, not exact-match.  A value like "female2"
      would match female (contains "female") and return 2.
    - The integer 1 or 2 is returned (not the string "1" or "2").

Known limitations:
    - Substring matching means unusual compound values could produce
      unexpected results (e.g. a value containing both "ז" and "נ" would
      match female first and return 2).
    - There is no status message returned by this engine; status text is
      added by the standardization pipeline layer above.
"""

from typing import Any


class GenderEngine:
    """
    Pure business logic for gender standardization.

    Normalizes gender values from various representations (Hebrew, English, numeric)
    to standardized codes: 1 (male) or 2 (female).

    This class replicates the exact behavior of the VBA NormalizeGenderValue function.
    """

    # Female patterns (case-insensitive substring match)
    FEMALE_PATTERNS = {"2", "female", "נ", "אישה", "בת", "f", "נקבה", "girl", "woman"}

    # Male patterns (case-insensitive substring match).
    # Checked only after female patterns fail so that "female" is never
    # accidentally matched by the "m" inside it.
    MALE_PATTERNS = {"1", "male", "ז", "זכר", "בן", "m", "man", "boy"}

    def normalize_gender(self, value: Any):
        """
        Normalize gender value to 1 (male), 2 (female), or "" (unrecognized).

        Algorithm:
        1. Convert value to string and trim whitespace.
        2. Convert to lowercase for case-insensitive matching.
        3. If empty, return 1 (male) — caller (pipeline) already short-circuits
           None/whitespace-only before reaching this method.
        4. If value contains any female pattern, return 2 (female).
        5. If value contains any male pattern, return 1 (male).
        6. Otherwise return "" — the value is not a recognized gender code and
           must not be copied as-is into the corrected field.

        Args:
            value: The gender value to normalize (can be string, int, or None)

        Returns:
            int 1 for male, int 2 for female, or "" for unrecognized values.

        Examples:
            >>> engine = GenderEngine()
            >>> engine.normalize_gender("2")
            2
            >>> engine.normalize_gender("female")
            2
            >>> engine.normalize_gender("נ")
            2
            >>> engine.normalize_gender("1")
            1
            >>> engine.normalize_gender("male")
            1
            >>> engine.normalize_gender("ז")
            1
            >>> engine.normalize_gender("8")
            ''
            >>> engine.normalize_gender("xyz")
            ''
        """
        # Convert to string and handle None/empty values
        if value is None:
            return 1

        # Convert to string, trim, and lowercase
        value_str = str(value).strip().lower()

        # Empty values default to male (pipeline short-circuits before this,
        # but keep the guard for direct callers)
        if not value_str:
            return 1

        # Check female patterns first
        for pattern in self.FEMALE_PATTERNS:
            if pattern.lower() in value_str:
                return 2

        # Check male patterns
        for pattern in self.MALE_PATTERNS:
            if pattern.lower() in value_str:
                return 1

        # Unrecognized value — return empty string, never copy raw value
        return ""
