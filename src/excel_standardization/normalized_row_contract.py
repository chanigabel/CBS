"""Shared contract helpers for normalized workbook rows.

The active pipeline keeps source values in their original fields and writes
standardized values only to ``*_corrected`` fields. These helpers centralize
the low-risk selection rules used by validation, UI preparation, and export.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Iterable, Mapping

SOURCE_FIELDS = {
    "first_name",
    "last_name",
    "father_name",
    "gender",
    "id_number",
    "passport",
    "birth_date",
    "birth_year",
    "birth_month",
    "birth_day",
    "entry_date",
    "entry_year",
    "entry_month",
    "entry_day",
    "MosadID",
    "SugMosad",
    "MisparDiraBeMosad",
}

CORRECTED_FIELDS = {
    "first_name_corrected",
    "last_name_corrected",
    "father_name_corrected",
    "gender_corrected",
    "id_number_corrected",
    "passport_corrected",
    "birth_date_corrected",
    "birth_year_corrected",
    "birth_month_corrected",
    "birth_day_corrected",
    "entry_date_corrected",
    "entry_year_corrected",
    "entry_month_corrected",
    "entry_day_corrected",
}

STATUS_FIELDS = {
    "gender_status",
    "identifier_status",
    "birth_date_status",
    "entry_date_status",
    "_validation_status",
    "_validation_ok",
}


@dataclass(frozen=True)
class NormalizedFieldGroup:
    """Source/corrected/status grouping metadata for a logical field family."""

    name: str
    source_fields: tuple[str, ...]
    corrected_fields: tuple[str, ...]
    status_field: str | None = None


NAME_FIELD_GROUP = NormalizedFieldGroup(
    name="name",
    source_fields=("first_name", "last_name", "father_name"),
    corrected_fields=("first_name_corrected", "last_name_corrected", "father_name_corrected"),
)

IDENTIFIER_FIELD_GROUP = NormalizedFieldGroup(
    name="identifier",
    source_fields=("id_number", "passport"),
    corrected_fields=("id_number_corrected", "passport_corrected"),
    status_field="identifier_status",
)

DATE_FIELD_GROUPS = (
    NormalizedFieldGroup(
        name="birth_date",
        source_fields=("birth_year", "birth_month", "birth_day", "birth_date"),
        corrected_fields=("birth_year_corrected", "birth_month_corrected", "birth_day_corrected"),
        status_field="birth_date_status",
    ),
    NormalizedFieldGroup(
        name="entry_date",
        source_fields=("entry_year", "entry_month", "entry_day", "entry_date"),
        corrected_fields=("entry_year_corrected", "entry_month_corrected", "entry_day_corrected"),
        status_field="entry_date_status",
    ),
)

GRID_GROUPS = (NAME_FIELD_GROUP, IDENTIFIER_FIELD_GROUP, *DATE_FIELD_GROUPS)

EXPORT_FIELD_TO_SOURCE: Dict[str, str] = {
    "MosadID": "MosadID",
    "SugMosad": "SugMosad",
    "MisparDiraBeMosad": "MisparDiraBeMosad",
    "ShemPrati": "first_name_corrected",
    "ShemMishpaha": "last_name_corrected",
    "ShemHaAv": "father_name_corrected",
    "MisparZehut": "id_number_corrected",
    "Darkon": "passport_corrected",
    "Min": "gender_corrected",
    "ShnatLida": "birth_year_corrected",
    "HodeshLida": "birth_month_corrected",
    "YomLida": "birth_day_corrected",
    "shnatknisa": "entry_year_corrected",
    "ShnatKnisa": "entry_year_corrected",
    "Hodeshknisa": "entry_month_corrected",
    "HodeshKnisa": "entry_month_corrected",
    "YomKnisa": "entry_day_corrected",
}


def is_blank(value: Any) -> bool:
    """Return True for None or whitespace-only values."""
    return value is None or str(value).strip() == ""


def corrected_field_name(source_field: str) -> str:
    """Return the conventional corrected field name for a source field."""
    return f"{source_field}_corrected"


def select_first_present(row: Mapping[str, Any], keys: Iterable[str]) -> Any:
    """Return the first non-blank value from ``keys`` in ``row``."""
    for key in keys:
        value = row.get(key)
        if not is_blank(value):
            return value
    return None


def select_corrected_or_original(row: Mapping[str, Any], source_field: str) -> Any:
    """Prefer a non-blank corrected value, then fall back to the source value."""
    return select_first_present(
        row,
        (corrected_field_name(source_field), source_field),
    )


def validation_source_value(row: Mapping[str, Any], source_field: str) -> Any:
    """Return the authoritative validation value for a normalized source field."""
    return select_corrected_or_original(row, source_field)


def select_corrected_only(row: Mapping[str, Any], corrected_field: str) -> Any:
    """Return a corrected value for corrected-only export, or ``""``."""
    value = row.get(corrected_field)
    return "" if is_blank(value) else value


def export_source_value(row: Mapping[str, Any], corrected_field: str) -> Any:
    """Return the approved export value for a corrected-only field."""
    return select_corrected_only(row, corrected_field)


def build_grid_group_maps():
    """Return helper maps used by the backend grid payload builder."""
    status_by_group = {
        group.name: group.status_field
        for group in GRID_GROUPS
        if group.status_field is not None
    }
    source_to_corrected = {
        source: corrected_field_name(source)
        for group in GRID_GROUPS
        for source in group.source_fields
        if not source.endswith("_date")
    }
    source_to_corrected.update({
        "birth_date": "birth_date_corrected",
        "entry_date": "entry_date_corrected",
    })
    return status_by_group, source_to_corrected


def build_standard_export_row(
    row: Mapping[str, Any],
    *,
    include_dira: bool = True,
    allow_mosad_fields: bool = True,
) -> Dict[str, Any]:
    """Map a normalized row to the canonical export field names.

    This intentionally preserves the approved corrected-only export policy for
    standardized fields.  Original values are not used as export fallback for
    names, gender, identifiers, or date components.
    """
    mapped: Dict[str, Any] = {}
    for export_field, source_field in EXPORT_FIELD_TO_SOURCE.items():
        if export_field == "MisparDiraBeMosad" and not include_dira:
            continue
        if export_field in {"MosadID", "SugMosad", "MisparDiraBeMosad"}:
            mapped[export_field] = (
                select_first_present(row, (source_field, source_field.lower()))
                if allow_mosad_fields
                else ""
            ) or ""
        else:
            mapped[export_field] = export_source_value(row, source_field)
    return mapped
