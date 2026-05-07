"""Gender normalization rules.

The engine maps common Hebrew, English, and numeric inputs to the canonical
codes 1 and 2. Unrecognized values return an empty string so the caller can
flag them.
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
            return ""

        # Convert to string, trim, and lowercase
        value_str = str(value).strip().lower()

        # Empty values default to male (pipeline short-circuits before this,
        # but keep the guard for direct callers)
        if not value_str:
            return ""

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
