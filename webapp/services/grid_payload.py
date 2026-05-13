"""Build the UI grid payload from an in-memory SheetDataset.

The workbook service keeps the public API compatible but delegates the
row-filtering, column-ordering, derived-column injection, and UI shaping to
this module so the logic stays isolated and testable.
"""

from __future__ import annotations

from typing import Any, Iterable, Mapping, Sequence

from webapp.models.responses import SheetDataResponse
from webapp.services.derived_columns import apply_derived_columns
from webapp.services.export_validation import row_has_visible_original_values, row_is_numeric_helper_row
from src.excel_standardization.normalized_row_contract import DATE_FIELD_GROUPS

_KEEP_INTERNAL = {"_row_uid", "_validation_status"}
_STATUS_GROUPS = {
    "gender_status": {"gender"},
    "identifier_status": {"id_number", "passport"},
    "birth_date_status": {"birth_year", "birth_month", "birth_day", "birth_date"},
    "entry_date_status": {"entry_year", "entry_month", "entry_day", "entry_date"},
}
_DATE_STRUCTURED_FALLBACK = {
    "birth_date_corrected": ["birth_day_corrected", "birth_month_corrected", "birth_year_corrected"],
    "entry_date_corrected": ["entry_day_corrected", "entry_month_corrected", "entry_year_corrected"],
}
_DATE_GROUPS = list(DATE_FIELD_GROUPS)


def _visible_row_copy(row: Mapping[str, Any]) -> dict[str, Any]:
    return {k: v for k, v in row.items() if not k.startswith("_standardization")}


def _all_row_keys(rows: Sequence[Mapping[str, Any]]) -> list[str]:
    seen: set[str] = set()
    keys: list[str] = []
    for row in rows:
        for key in row.keys():
            if key not in seen and not key.startswith("_"):
                seen.add(key)
                keys.append(key)
    return keys


def _build_anchor_to_status(original_fields: Sequence[str], seen: Iterable[str]) -> dict[str, str]:
    seen_set = set(seen)
    anchor_to_status: dict[str, str] = {}
    for status_key, group_members in _STATUS_GROUPS.items():
        if status_key not in seen_set:
            continue
        anchor_orig = None
        for field in original_fields:
            if field in group_members:
                anchor_orig = field
        if anchor_orig is None:
            continue
        anchor_corrected = f"{anchor_orig}_corrected"
        if anchor_corrected not in seen_set and anchor_corrected in _DATE_STRUCTURED_FALLBACK:
            for fallback in _DATE_STRUCTURED_FALLBACK[anchor_corrected]:
                if fallback in seen_set:
                    anchor_corrected = fallback
                    break
        anchor_to_status[anchor_corrected] = status_key
    return anchor_to_status


def _build_display_columns(original_fields: Sequence[str], rows: Sequence[Mapping[str, Any]]) -> list[str]:
    seen = set(_all_row_keys(rows))

    generated_identifier_corrected = {"passport_corrected"} if "passport" not in original_fields else set()
    display_columns: list[str] = []
    placed: set[str] = set()
    date_groups_emitted: set[str] = set()
    anchor_to_status = _build_anchor_to_status(original_fields, seen)

    for orig in original_fields:
        if orig in generated_identifier_corrected:
            continue

        owning_group = None
        for dg in _DATE_GROUPS:
            if orig in dg.source_fields:
                owning_group = dg
                break

        if owning_group is not None:
            status_key = owning_group.status_field
            if status_key and status_key not in date_groups_emitted:
                date_groups_emitted.add(status_key)
                for src in original_fields:
                    if src in owning_group.source_fields and src not in placed:
                        display_columns.append(src)
                        placed.add(src)
                for cf in owning_group.corrected_fields:
                    if cf in seen and cf not in placed:
                        display_columns.append(cf)
                        placed.add(cf)
                if status_key in seen and status_key not in placed:
                    display_columns.append(status_key)
                    placed.add(status_key)
            continue

        if orig not in placed:
            display_columns.append(orig)
            placed.add(orig)

        corrected = f"{orig}_corrected"
        if corrected in seen and corrected not in placed:
            display_columns.append(corrected)
            placed.add(corrected)

        if orig == "id_number":
            for generated_corrected in generated_identifier_corrected:
                if generated_corrected in seen and generated_corrected not in placed:
                    display_columns.append(generated_corrected)
                    placed.add(generated_corrected)

        status_key = anchor_to_status.get(corrected)
        if status_key and status_key in seen and status_key not in placed:
            display_columns.append(status_key)
            placed.add(status_key)

    for key in _all_row_keys(rows):
        if key not in placed:
            display_columns.append(key)
            placed.add(key)

    if "_validation_status" not in placed:
        if any("_validation_status" in row for row in rows):
            display_columns.append("_validation_status")
            placed.add("_validation_status")

    return display_columns


def _filter_visible_rows(rows: Sequence[Mapping[str, Any]], original_field_set: set[str]) -> list[dict[str, Any]]:
    clean_rows = [row for row in (_visible_row_copy(row) for row in rows) if row_has_visible_original_values(row, original_field_set)]
    if clean_rows and row_is_numeric_helper_row(clean_rows[0], original_field_set):
        clean_rows = clean_rows[1:]
    return clean_rows


def build_sheet_grid_payload(
    sheet,
    *,
    session_mosad_id: str = "",
    active_mosad_type: str | None = None,
    metadata_mosad_id: str | None = None,
) -> SheetDataResponse:
    """Build the sheet payload exactly as the UI expects it."""
    original_fields = list(sheet.field_names)
    original_field_set = set(original_fields)
    clean_rows = _filter_visible_rows(sheet.rows, original_field_set)
    display_columns = _build_display_columns(original_fields, clean_rows)

    meta_mosad_id = session_mosad_id or metadata_mosad_id
    clean_rows, display_columns = apply_derived_columns(
        rows=clean_rows,
        field_names=original_fields,
        display_columns=display_columns,
        meta_mosad_id=meta_mosad_id,
    )

    sug_mosad_in_rows = any(row.get("SugMosad") for row in clean_rows)
    if active_mosad_type:
        for row in clean_rows:
            if not row.get("SugMosad"):
                row["SugMosad"] = active_mosad_type
        sug_mosad_in_rows = True
    if sug_mosad_in_rows:
        if "SugMosad" in display_columns:
            display_columns.remove("SugMosad")
        try:
            insert_pos = display_columns.index("MosadID") + 1
        except ValueError:
            insert_pos = 1
        display_columns.insert(insert_pos, "SugMosad")

    return SheetDataResponse(
        sheet_name=sheet.sheet_name,
        field_names=display_columns,
        rows=clean_rows,
    )
