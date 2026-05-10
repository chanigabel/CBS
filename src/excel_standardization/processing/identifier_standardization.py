"""Identifier standardization helpers for the processing pipeline."""

from __future__ import annotations

import logging
from typing import Any, List, Optional

from ..data_types import JsonRow

logger = logging.getLogger(__name__)


# הפונקציה מנרמלת מזהים בשורה ומפרידה בין תעודת זהות, דרכון וסטטוס שגיאה.
def apply_identifier_standardization(
    pipeline: Any,
    json_row: JsonRow,
    row_number: Optional[int] = None,
) -> List[str]:
    """Apply IdentifierEngine to identifier fields in the row."""
    failed_fields: List[str] = []

    id_value = json_row.get("id_number")
    passport_value = json_row.get("passport")

    if "id_number" not in json_row and "passport" not in json_row:
        return failed_fields

    if (id_value is None or id_value == "") and (passport_value is None or passport_value == ""):
        if "id_number" in json_row:
            json_row["id_number_corrected"] = id_value
        if "passport" in json_row:
            json_row["passport_corrected"] = passport_value
        json_row["identifier_status"] ="חסר מזהים"
        return failed_fields

    try:
        result = pipeline.identifier_engine.normalize_identifiers(id_value, passport_value)

        if "id_number" in json_row:
            json_row["id_number_corrected"] = result.corrected_id
        if "passport" in json_row or result.corrected_passport:
            json_row["passport_corrected"] = result.corrected_passport
        json_row["identifier_status"] = result.status_text

    except Exception as e:
        if "id_number" in json_row:
            json_row["id_number_corrected"] = id_value
            failed_fields.append("id_number")
        if "passport" in json_row:
            json_row["passport_corrected"] = passport_value
            failed_fields.append("passport")
        json_row["identifier_status"] = ""

        row_info = f"row {row_number}" if row_number is not None else "unknown row"
        logger.error(
            f"Identifier standardization failed for fields 'id_number'/'passport' at {row_info}: {str(e)}. "
            f"Original values: id_number='{id_value}', passport='{passport_value}'"
        )

    return failed_fields
