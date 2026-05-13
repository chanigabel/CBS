"""Gender standardization helpers for the processing pipeline."""

from __future__ import annotations

import logging
from typing import Any, List, Optional

from ..data_types import JsonRow

logger = logging.getLogger(__name__)


# הפונקציה מנרמלת ערך מגדר בשורה ומוסיפה סטטוס שמוצג בהמשך ב־UI ובדוחות.
def apply_gender_standardization(
    pipeline: Any,
    json_row: JsonRow,
    row_number: Optional[int] = None,
) -> List[str]:
    """Apply GenderEngine to gender field in the row."""
    failed_fields: List[str] = []

    if "gender" in json_row:
        original = json_row["gender"]

        if original is None:
            json_row["gender_corrected"] = original
            return failed_fields

        if str(original).strip() == "":
            json_row["gender_corrected"] = ""
            return failed_fields

        try:
            corrected = pipeline.gender_engine.normalize_gender(original)
            json_row["gender_corrected"] = corrected
            if corrected == "":
                json_row["gender_status"] = "קוד מין לא תקין - חייב להיות 1 (זכר) או 2 (נקבה)"
                json_row.setdefault("gender_status", "")
        except Exception as e:
            json_row["gender_corrected"] = original
            failed_fields.append("gender")

            row_info = f"row {row_number}" if row_number is not None else "unknown row"
            logger.error(
                f"Gender standardization failed for field 'gender' at {row_info}: {str(e)}. "
                f"Original value: '{original}'"
            )

    return failed_fields
