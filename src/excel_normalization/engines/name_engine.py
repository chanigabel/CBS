"""NameEngine — name field normalization rules.

Purpose:
    Applies name cleaning and last-name removal to first name, last name,
    and father name fields.  Delegates all character-level cleaning to
    TextProcessor.clean_name() and adds field-level logic on top.

Implemented rules:

    1. Basic name cleaning (all name fields)
       Every name value is passed through TextProcessor.clean_name() before
       any field-specific logic runs.  This covers: digit removal, diacritic
       removal, hyphen-to-space conversion, unwanted token removal, and space
       normalisation.  See text_processor.py for the full pipeline.
       Example:
           'יוסי ז"ל'   → "יוסי"
           "  כהן  "    → "כהן"
           "אבר9הם"     → "אברהם"

    2. Last-name removal from father name (two-stage)
       When a detected pattern shows that the last name is consistently
       embedded in the father name field, it is removed using two stages:

       Stage A — substring removal (word-boundary aware):
           Runs only when the last name appears as a substring of the father
           name.  Uses space-padded matching so partial word matches are
           avoided.
           Example:
               father="כהן אברהם", last="כהן"  → "אברהם"
               father="אברהם כהן", last="כהן"  → "אברהם"

       Stage B — positional fallback (runs only if Stage A made no change):
           If the last name is not a substring, removes the first or last
           token depending on the detected pattern.
           REMOVE_FIRST: removes the first word.
               father="כהן אברהם", pattern=REMOVE_FIRST → "אברהם"
           REMOVE_LAST: removes the last word.
               father="אברהם כהן", pattern=REMOVE_LAST  → "אברהם"

       Single-word guard: if the father name is a single word, neither stage
       modifies it.

       NONE pattern: no removal is attempted.

    3. Last-name removal from first name (two-stage, same logic)
       Identical two-stage logic applied to the first name field when the
       pattern detection finds the last name consistently embedded there.
       Example:
           first="כהן יוסי", last="כהן", pattern=REMOVE_FIRST → "יוסי"

    4. Pattern detection (father name and first name)
       Samples up to 5 rows to decide whether the last name is embedded in
       the father/first name field and in which position.
       Decision rules (applied to the sample):
       - If the last name appears in fewer than 3 out of 5 rows → NONE
         (no removal).
       - If it appears at the start in ≥ 3 rows → REMOVE_FIRST.
       - If it appears at the end in ≥ 3 rows → REMOVE_LAST.
       - Otherwise → NONE.

Important notes:
    - NameEngine does not read or write Excel files.
    - It does not call any web or I/O layer.
    - All character-level cleaning is delegated to TextProcessor.
    - The engine operates on already-extracted string values.
    - normalize_name() is a direct alias for TextProcessor.clean_name().

Known limitations:
    - Pattern detection samples only the first 5 rows.  If the pattern is
      inconsistent across the dataset, NONE is returned and no removal occurs.
    - Stage A uses simple substring matching with space padding; it does not
      handle hyphenated compound last names (e.g. "בן-דוד") as a single unit
      after the hyphen has been converted to a space by clean_name().
    - The two-stage removal is applied uniformly to all rows once the pattern
      is detected from the sample; there is no per-row confidence check.
"""

import logging
from typing import List, Sequence

from .text_processor import TextProcessor
from ..data_types import FatherNamePattern

logger = logging.getLogger(__name__)


class NameEngine:
    def __init__(self, text_processor: TextProcessor):
        self.text_processor = text_processor

    # ------------------------------------------------------------------
    # בסיס
    # ------------------------------------------------------------------

    def normalize_name(self, name) -> str:
        return self.text_processor.clean_name(name)

    def normalize_names(self, input_data: Sequence[Sequence]) -> List[List[str]]:
        result: List[List[str]] = []

        for row in input_data:
            value = row[0] if row else ""
            result.append([self.normalize_name(value)])

        return result

    # ------------------------------------------------------------------
    # שם פרטי (🔥 ללא pattern)
    # ------------------------------------------------------------------

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

    def remove_last_name_from_first_name(
        self,
        first_name: str,
        last_name: str,
        pattern: FatherNamePattern = FatherNamePattern.NONE,
    ) -> str:
        """Remove the last name from a first name field using two-stage logic.

        Stage A: substring removal (word-boundary aware).
        Stage B: positional fallback — runs ONLY if Stage A did not change the value.

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

        # Single-word first name — never modify
        if len(first_name.split()) == 1:
            return first_name

        # ------------------------------------------------------------------
        # Stage A: substring removal
        # Only runs when the last name actually appears as a substring.
        # When it doesn't, Stage A is a guaranteed no-op, so fall through
        # directly to Stage B.
        # ------------------------------------------------------------------
        if last_name in first_name:
            after_stage_a = self.text_processor.remove_substring(first_name, last_name)

            if not after_stage_a.strip():
                return ""

            # Stage A changed the value → stop, do NOT run Stage B.
            if after_stage_a != first_name:
                return after_stage_a
        # else: last_name not a substring → Stage A would make no change,
        #       fall through directly to Stage B.

        # ------------------------------------------------------------------
        # Stage B: positional fallback
        # Runs when Stage A made no change (either because the substring was
        # absent, or because remove_substring left the value identical).
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

            if pattern != FatherNamePattern.NONE and last:
                father = self.remove_last_name_from_father(father, last, pattern)

            result.append([father])

        return result

    def remove_last_name_from_father(
        self,
        father_name: str,
        last_name: str,
        pattern: FatherNamePattern,
    ) -> str:
        """Remove the last name from a father name field using two-stage logic.

        Stage A: substring removal (word-boundary aware via remove_substring).
        Stage B: positional fallback — runs ONLY if Stage A did not change the value.

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

        # NONE pattern → never modify
        if pattern == FatherNamePattern.NONE:
            return father_name

        # ------------------------------------------------------------------
        # Stage A: substring removal
        # Only runs when the last name actually appears as a substring.
        # When it doesn't, Stage A is a guaranteed no-op, so skip straight
        # to Stage B.
        # ------------------------------------------------------------------
        if last_name in father_name:
            after_stage_a = self.text_processor.remove_substring(father_name, last_name)

            if not after_stage_a.strip():
                return ""

            # Stage A changed the value → stop, do NOT run Stage B.
            if after_stage_a != father_name:
                return after_stage_a
        # else: last_name not a substring → Stage A would make no change,
        #       fall through directly to Stage B.

        # ------------------------------------------------------------------
        # Stage B: positional fallback
        # Runs when Stage A made no change (either because the substring was
        # absent, or because remove_substring left the value identical).
        # ------------------------------------------------------------------
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
            father = self.text_processor.safe_to_string(father_sample[i][0]).strip()
            ln = self.text_processor.safe_to_string(last_name_sample[i][0]).strip()

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
            fn = self.text_processor.safe_to_string(first_name_sample[i][0]).strip()
            ln = self.text_processor.safe_to_string(last_name_sample[i][0]).strip()

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