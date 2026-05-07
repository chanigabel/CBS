"""Name standardization helpers for the processing pipeline."""

from __future__ import annotations

import logging
from typing import List, Optional, Any

from ..data_types import FatherNamePattern, JsonRow

logger = logging.getLogger(__name__)


def apply_name_standardization(
    pipeline: Any,
    json_row: JsonRow,
    row_number: Optional[int] = None,
) -> List[str]:
    """Apply NameEngine to name fields in the row."""
    failed_fields: List[str] = []

    try:
        if "last_name" in json_row:
            original = json_row["last_name"]
            if original is None or original == "":
                json_row["last_name_corrected"] = original
            else:
                json_row["last_name_corrected"] = pipeline.name_engine.normalize_name(str(original))

        cleaned_last = ""
        if "last_name" in json_row:
            raw_last = json_row.get("last_name")
            if raw_last:
                cleaned_last = pipeline.name_engine.normalize_name(str(raw_last))

        if "first_name" in json_row:
            original = json_row["first_name"]
            if original is None or original == "":
                json_row["first_name_corrected"] = original
            else:
                cleaned = pipeline.name_engine.normalize_name(str(original))
                if cleaned_last:
                    pattern = getattr(pipeline, "_first_name_pattern", FatherNamePattern.NONE)
                    cleaned = pipeline.name_engine.remove_last_name_from_first_name(
                        cleaned, cleaned_last, pattern
                    )
                json_row["first_name_corrected"] = cleaned

        if "father_name" in json_row:
            original = json_row["father_name"]
            if original is None or original == "":
                json_row["father_name_corrected"] = original
            else:
                cleaned = pipeline.name_engine.normalize_name(str(original))
                if cleaned_last:
                    pattern = getattr(pipeline, "_father_name_pattern", FatherNamePattern.NONE)
                    cleaned = pipeline.name_engine.remove_last_name_from_father(
                        cleaned, cleaned_last, pattern
                    )
                json_row["father_name_corrected"] = cleaned

    except Exception as e:
        for field in ["first_name", "last_name", "father_name"]:
            if field in json_row and f"{field}_corrected" not in json_row:
                json_row[f"{field}_corrected"] = json_row[field]
                failed_fields.append(field)
        row_info = f"row {row_number}" if row_number is not None else "unknown row"
        logger.error(f"Name standardization failed at {row_info}: {e}")

    return failed_fields
