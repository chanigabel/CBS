"""Name standardization rules for first, last, and father names.

The engine delegates character cleanup to TextProcessor and applies the field-
level removal rules that normalize embedded last-name fragments.
"""

import logging
from typing import List, Sequence

from .text_processor import TextProcessor
from ..data_types import FatherNamePattern

logger = logging.getLogger(__name__)


# המנוע אחראי לניקוי שמות ולזיהוי דפוסי שם משפחה בשדות שם.
class NameEngine:
    engine_key = "name"
    display_name = "Name Engine"
    version = "1.0.0"
    description = "Standardizes first, last, and father name fields."
    supported_fields = ["first_name", "last_name", "father_name"]

    # הפונקציה מקבלת TextProcessor כדי לרכז את כל ניקויי הטקסט במקום אחד.
    def __init__(self, text_processor: TextProcessor):
        self.text_processor = text_processor

    # ------------------------------------------------------------------
    # בסיס
    # ------------------------------------------------------------------

    # הפונקציה מנרמלת שם יחיד דרך מנגנון ניקוי הטקסט.
    def normalize_name(self, name) -> str:
        return self.text_processor.clean_name(name)

    # הפונקציה מנרמלת מטריצת שמות עבור עיבוד batch והתאמה להתנהגות legacy.
    def normalize_names(self, input_data: Sequence[Sequence]) -> List[List[str]]:
        result: List[List[str]] = []

        for row in input_data:
            value = row[0] if row else ""
            result.append([self.normalize_name(value)])

        return result

    # ------------------------------------------------------------------
    # שם פרטי (🔥 ללא pattern)
    # ------------------------------------------------------------------

    # הפונקציה מנרמלת שמות פרטיים ומסירה שם משפחה כאשר זוהה דפוס מתאים.
    def normalize_first_names(
        self,
        first_name_data: Sequence[Sequence],
        last_name_data: Sequence[Sequence],
    ) -> List[List[str]]:

        rows = max(len(first_name_data), len(last_name_data))
        result: List[List[str]] = []

        # Detect whether the last name is embedded in the first name field
        # and in which position, so Stage B fallback uses the right pattern.
        pattern = self.detect_first_name_pattern(first_name_data, last_name_data)

        for i in range(rows):
            raw = first_name_data[i][0] if i < len(first_name_data) and first_name_data[i] else ""
            last_raw = last_name_data[i][0] if i < len(last_name_data) and last_name_data[i] else ""

            first = self.normalize_name(raw)
            last = self.normalize_name(last_raw) if last_raw else ""

            if last:
                first = self.remove_last_name_from_first_name(first, last, pattern)

            result.append([first])

        return result

    # הפונקציה מסירה שם משפחה מתוך שם פרטי לפי דפוס הגיליון שנבחר.
    def remove_last_name_from_first_name(
        self,
        first_name: str,
        last_name: str,
        pattern: FatherNamePattern = FatherNamePattern.NONE,
    ) -> str:
        """Remove the last name from a first name field using two-stage logic.

        Stage A: row-local substring removal (word-boundary aware).
        Stage B: dataset-level positional fallback when a pattern was detected
        and Stage A did not change the value.

        Args:
            first_name: Cleaned first name string.
            last_name:  Cleaned last name string.
            pattern:    Detected removal pattern (REMOVE_FIRST / REMOVE_LAST / NONE).
                        When NONE, Stage B is skipped entirely.

        Returns:
            First name with the last name removed, or the original if no change
            is warranted.
        """
        first_name = self.text_processor.safe_to_string(first_name).strip()
        last_name = self.text_processor.safe_to_string(last_name).strip()

        if not first_name or not last_name:
            return first_name

        # ------------------------------------------------------------------
        # Stage A: row-local substring removal.
        # ------------------------------------------------------------------
        if last_name in first_name:
            after_stage_a = self.text_processor.remove_substring(first_name, last_name)

            if not after_stage_a.strip():
                return ""

        # Stage A changed the value → stop, do NOT run Stage B.
            if after_stage_a != first_name:
                return after_stage_a

        # ------------------------------------------------------------------
        # Stage B: positional fallback
        # Runs when dataset-level detection found a consistent positional
        # pattern and Stage A did not change the value.
        # ------------------------------------------------------------------
        if pattern == FatherNamePattern.NONE:
            return first_name

        parts = first_name.split()

        # Only apply positional removal when at least 2 words remain
        if len(parts) < 2:
            return first_name

        if pattern == FatherNamePattern.REMOVE_FIRST:
            return " ".join(parts[1:])

        if pattern == FatherNamePattern.REMOVE_LAST:
            return " ".join(parts[:-1])

        return first_name

    # ------------------------------------------------------------------
    # שם אב (🔥 עם pattern)
    # ------------------------------------------------------------------

    # הפונקציה מנרמלת שמות אב ומחילה הסרת שם משפחה בהתאם לדפוס שזוהה.
    def normalize_father_names(
        self,
        father_data: Sequence[Sequence],
        last_name_data: Sequence[Sequence],
        pattern: FatherNamePattern,
    ) -> List[List[str]]:

        rows = max(len(father_data), len(last_name_data))
        result: List[List[str]] = []

        for i in range(rows):
            father_raw = father_data[i][0] if i < len(father_data) and father_data[i] else ""
            last_raw = last_name_data[i][0] if i < len(last_name_data) and last_name_data[i] else ""

            father = self.normalize_name(father_raw)
            last = self.normalize_name(last_raw) if last_raw else ""

            if last:
                father = self.remove_last_name_from_father(father, last, pattern)

            result.append([father])

        return result

    # הפונקציה מסירה שם משפחה משם האב לפי כללי ההתאמה של המערכת.
    def remove_last_name_from_father(
        self,
        father_name: str,
        last_name: str,
        pattern: FatherNamePattern,
    ) -> str:
        """Remove the last name from a father name field using two-stage logic.

        Stage A: row-local substring removal (word-boundary aware).
        Stage B: dataset-level positional fallback when a pattern was detected
        and Stage A did not change the value.

        Args:
            father_name: Cleaned father name string.
            last_name:   Cleaned last name string.
            pattern:     Detected removal pattern (REMOVE_FIRST / REMOVE_LAST / NONE).

        Returns:
            Father name with the last name removed, or the original if no change
            is warranted.
        """
        father_name = self.text_processor.safe_to_string(father_name).strip()
        last_name = self.text_processor.safe_to_string(last_name).strip()

        if not father_name or not last_name:
            return father_name

        # ------------------------------------------------------------------
        # Stage A: row-local substring removal.
        # ------------------------------------------------------------------
        if last_name in father_name:
            after_stage_a = self.text_processor.remove_substring(father_name, last_name)

            if not after_stage_a.strip():
                return ""

        # Stage A changed the value → stop, do NOT run Stage B.
            if after_stage_a != father_name:
                return after_stage_a
        # ------------------------------------------------------------------
        # Stage B: positional fallback
        # Runs when dataset-level detection found a consistent positional
        # pattern and Stage A did not change the value.
        # ------------------------------------------------------------------
        if pattern == FatherNamePattern.NONE:
            return father_name

        parts = father_name.split()

        # Only apply positional removal when at least 2 words remain
        if len(parts) < 2:
            return father_name

        if pattern == FatherNamePattern.REMOVE_FIRST:
            return " ".join(parts[1:])

        if pattern == FatherNamePattern.REMOVE_LAST:
            return " ".join(parts[:-1])

        return father_name

    # ------------------------------------------------------------------
    # זיהוי pattern
    # ------------------------------------------------------------------

    # הפונקציה מזהה אם שם האב כולל את שם המשפחה כדי להפעיל תיקון עקבי בגיליון.
    def detect_father_name_pattern(
        self,
        father_sample: Sequence[Sequence],
        last_name_sample: Sequence[Sequence],
    ) -> FatherNamePattern:

        sample_size = min(5, len(father_sample), len(last_name_sample))
        if sample_size <= 0:
            return FatherNamePattern.NONE

        contain = 0
        first = 0
        last = 0

        for i in range(sample_size):
            father = self.normalize_name(father_sample[i][0])
            ln = self.normalize_name(last_name_sample[i][0])

            if not father or not ln:
                continue

            if ln in father:
                contain += 1
                parts = father.split()

                if parts and parts[0] == ln:
                    first += 1
                if parts and parts[-1] == ln:
                    last += 1

        if contain < 3:
            return FatherNamePattern.NONE

        if first >= 3:
            return FatherNamePattern.REMOVE_FIRST

        if last >= 3:
            return FatherNamePattern.REMOVE_LAST

        return FatherNamePattern.NONE

    # הפונקציה מזהה אם השם הפרטי כולל שם משפחה כדי לקבוע אסטרטגיית הסרה.
    def detect_first_name_pattern(
        self,
        first_name_sample: Sequence[Sequence],
        last_name_sample: Sequence[Sequence],
    ) -> FatherNamePattern:
        """Detect whether the last name is embedded in the first name field.

        Mirrors detect_father_name_pattern: samples up to 5 rows, counts how
        often the last name appears in the first name field and whether it tends
        to be at the start or end.

        Returns:
            FatherNamePattern.REMOVE_FIRST  — last name is usually the first token
            FatherNamePattern.REMOVE_LAST   — last name is usually the last token
            FatherNamePattern.NONE          — last name not consistently embedded
        """
        sample_size = min(5, len(first_name_sample), len(last_name_sample))
        if sample_size <= 0:
            return FatherNamePattern.NONE

        contain = 0
        first_pos = 0
        last_pos = 0

        for i in range(sample_size):
            fn = self.normalize_name(first_name_sample[i][0])
            ln = self.normalize_name(last_name_sample[i][0])

            if not fn or not ln:
                continue

            if ln in fn:
                contain += 1
                parts = fn.split()

                if parts and parts[0] == ln:
                    first_pos += 1
                if parts and parts[-1] == ln:
                    last_pos += 1

        if contain < 3:
            return FatherNamePattern.NONE

        if first_pos >= 3:
            return FatherNamePattern.REMOVE_FIRST

        if last_pos >= 3:
            return FatherNamePattern.REMOVE_LAST

        return FatherNamePattern.NONE
