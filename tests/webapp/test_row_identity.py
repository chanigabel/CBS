from src.excel_standardization.data_types import SheetDataset
from webapp.services.grid_payload import build_sheet_grid_payload
from webapp.services.row_identity import ensure_sheet_row_uids, find_row_by_uid


def test_grid_payload_creates_missing_row_uid_once_and_persists_it():
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["name"],
        rows=[{"name": "Alice"}, {"name": "Bob"}],
    )

    first = build_sheet_grid_payload(sheet)
    second = build_sheet_grid_payload(sheet)

    first_uids = [row["_row_uid"] for row in first.rows]
    second_uids = [row["_row_uid"] for row in second.rows]
    assert first_uids == second_uids
    assert [row["_row_uid"] for row in sheet.rows] == first_uids


def test_row_lookup_supports_legacy_row_uid_alias_and_normalizes_to_canonical():
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["name"],
        rows=[{"row_uid": "legacy-1", "name": "Alice"}],
    )

    ensure_sheet_row_uids(sheet)
    found = find_row_by_uid(sheet, "legacy-1")

    assert found is not None
    assert sheet.rows[0]["_row_uid"] == "legacy-1"
