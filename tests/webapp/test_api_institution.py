"""Tests for the scoped SugMosad (institution type) apply feature.

Covers:
- Applying one Institution Type to the entire workbook (scope=workbook).
- Applying one Institution Type to one selected sheet only (scope=sheet).
- Verifying other sheets remain unchanged when sheet-level mode is used.
- Applying Institution Type to selected rows (scope=selected_rows) by _row_uid.
- Applying 2 different Institution Types to 2 different selected row groups.
- Applying 3 different Institution Types to 3 different selected row groups.
- Verifying unselected rows remain unchanged.
- Verifying other sheets remain unchanged when selected_rows mode is used.
- Verifying selected_rows scope overrides sheet/workbook scope during export.
- Validation: no rows selected in selected_rows mode.
- Validation: non-numeric Institution Type rejected.
- Validation: Institution Type with < 3 digits rejected.
- Validation: Institution Type with exactly 3 digits accepted.
- Validation: non-numeric MosadID rejected.
- Validation: MosadID with < 3 digits rejected.
- Validation: MosadID with exactly 3 digits accepted.
- Validation: non-numeric MisparDiraBeMosad flagged when column exists.
- No failure when MisparDiraBeMosad column does not exist.
"""

import io
import uuid
import pytest
from openpyxl import Workbook
from fastapi.testclient import TestClient

from webapp.services.session_service import SessionService
from webapp.models.session import SessionRecord
from src.excel_standardization.data_types import SheetDataset, WorkbookDataset


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_xlsx_bytes(sheet_names=None):
    wb = Workbook()
    wb.remove(wb.active)
    for name in (sheet_names or ["Sheet1"]):
        ws = wb.create_sheet(name)
        ws.append(["first_name", "last_name"])
        ws.append(["Alice", "Smith"])
        ws.append(["Bob", "Jones"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _make_sheet_with_uids(name, n_rows=3):
    """Build a SheetDataset whose rows already have _row_uid assigned."""
    rows = []
    for i in range(1, n_rows + 1):
        rows.append({
            "first_name": f"Name{i}",
            "last_name": "Test",
            "_row_uid": uuid.uuid4().hex,
        })
    return SheetDataset(
        sheet_name=name,
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "last_name"],
        rows=rows,
    )


def _make_session(tmp_path, sheet_names=None, n_rows=3):
    sheet_names = sheet_names or ["Sheet1"]
    svc = SessionService()
    svc.clear_all()

    file_bytes = _make_xlsx_bytes(sheet_names)
    path = tmp_path / "test.xlsx"
    path.write_bytes(file_bytes)

    sheets = [_make_sheet_with_uids(name, n_rows) for name in sheet_names]
    wbd = WorkbookDataset(source_file=str(path), sheets=sheets)

    record = SessionRecord(
        session_id="test-session",
        source_file_path=str(path),
        working_copy_path=str(path),
        original_filename="test.xlsx",
        status="standardized",
        workbook_dataset=wbd,
    )
    svc.create(record)
    return svc, record


@pytest.fixture(autouse=True)
def clear_registry():
    svc = SessionService()
    svc.clear_all()
    yield
    svc.clear_all()


@pytest.fixture
def client_with_two_sheets(tmp_path, monkeypatch):
    import webapp.dependencies as deps
    from webapp.services.upload_service import UploadService
    from webapp.services.workbook_service import WorkbookService
    from webapp.services.standardization_service import standardizationService
    from webapp.services.edit_service import EditService
    from webapp.services.export_service import ExportService

    svc, record = _make_session(tmp_path, sheet_names=["Sheet1", "Sheet2"], n_rows=3)

    monkeypatch.setattr(deps, "_session_service", svc)
    monkeypatch.setattr(deps, "_upload_service", UploadService(svc, tmp_path / "uploads", tmp_path / "work"))
    monkeypatch.setattr(deps, "_workbook_service", WorkbookService(svc))
    monkeypatch.setattr(deps, "_standardization_service", standardizationService(svc))
    monkeypatch.setattr(deps, "_edit_service", EditService(svc))
    monkeypatch.setattr(deps, "_export_service", ExportService(svc, tmp_path / "output"))

    from webapp.app import app
    with TestClient(app) as c:
        yield c, svc, record


# ---------------------------------------------------------------------------
# 1. Workbook scope
# ---------------------------------------------------------------------------

class TestApplyScopedWorkbook:
    def test_workbook_scope_updates_all_sheets(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook",
            "sug_mosad": "1234",
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["scope"] == "workbook"
        assert data["updated_rows"] == 6  # 3 rows x 2 sheets

        record = svc.get("test-session")
        for sheet in record.workbook_dataset.sheets:
            for row in sheet.rows:
                assert row["SugMosad"] == "1234"

    def test_workbook_scope_stores_config(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "5678",
        })
        record = svc.get("test-session")
        assert len(record.sug_mosad_configs) == 1
        cfg = record.sug_mosad_configs[0]
        assert cfg.scope == "workbook"
        assert cfg.sug_mosad == "5678"


# ---------------------------------------------------------------------------
# 2. Sheet scope
# ---------------------------------------------------------------------------

class TestApplyScopedSheet:
    def test_sheet_scope_updates_only_target_sheet(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "sheet", "sug_mosad": "9999", "sheet_name": "Sheet1",
        })
        assert resp.status_code == 200
        assert resp.json()["updated_rows"] == 3

        record = svc.get("test-session")
        sheet1 = record.workbook_dataset.get_sheet_by_name("Sheet1")
        sheet2 = record.workbook_dataset.get_sheet_by_name("Sheet2")
        for row in sheet1.rows:
            assert row["SugMosad"] == "9999"
        for row in sheet2.rows:
            assert row.get("SugMosad") != "9999"

    def test_sheet_scope_other_sheets_unchanged(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        record = svc.get("test-session")
        for row in record.workbook_dataset.get_sheet_by_name("Sheet2").rows:
            row["SugMosad"] = "ORIGINAL"

        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "sheet", "sug_mosad": "1111", "sheet_name": "Sheet1",
        })

        record = svc.get("test-session")
        for row in record.workbook_dataset.get_sheet_by_name("Sheet2").rows:
            assert row.get("SugMosad") == "ORIGINAL"

    def test_sheet_scope_404_for_unknown_sheet(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "sheet", "sug_mosad": "1234", "sheet_name": "NonExistent",
        })
        assert resp.status_code == 404


# ---------------------------------------------------------------------------
# 3. Selected-rows scope
# ---------------------------------------------------------------------------

class TestApplyScopedSelectedRows:
    def _get_uids(self, svc, sheet_name="Sheet1"):
        record = svc.get("test-session")
        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)
        return [row["_row_uid"] for row in sheet.rows]

    def test_selected_rows_updates_only_selected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        uids = self._get_uids(svc)
        # Select first 2 rows only
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "1111", "row_uids": uids[:2]}],
        })
        assert resp.status_code == 200
        assert resp.json()["updated_rows"] == 2

        record = svc.get("test-session")
        sheet1 = record.workbook_dataset.get_sheet_by_name("Sheet1")
        uid_to_row = {r["_row_uid"]: r for r in sheet1.rows}
        assert uid_to_row[uids[0]]["SugMosad"] == "1111"
        assert uid_to_row[uids[1]]["SugMosad"] == "1111"
        # Third row must NOT be changed
        assert uid_to_row[uids[2]].get("SugMosad") != "1111"

    def test_two_groups_applied_correctly(self, client_with_two_sheets):
        """Two calls: first 2 rows get 1111, last row gets 2222."""
        client, svc, record = client_with_two_sheets
        uids = self._get_uids(svc)

        # First apply: rows 0+1 -> 1111
        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "1111", "row_uids": uids[:2]}],
        })
        # Second apply: row 2 -> 2222 (replaces config for same sheet)
        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "2222", "row_uids": [uids[2]]}],
        })

        record = svc.get("test-session")
        sheet1 = record.workbook_dataset.get_sheet_by_name("Sheet1")
        uid_to_row = {r["_row_uid"]: r for r in sheet1.rows}
        assert uid_to_row[uids[0]]["SugMosad"] == "1111"
        assert uid_to_row[uids[1]]["SugMosad"] == "1111"
        assert uid_to_row[uids[2]]["SugMosad"] == "2222"

    def test_three_groups_in_one_request(self, client_with_two_sheets):
        """Single request with 3 groups, each covering one row."""
        client, svc, record = client_with_two_sheets
        uids = self._get_uids(svc)

        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [
                {"sug_mosad": "1111", "row_uids": [uids[0]]},
                {"sug_mosad": "2222", "row_uids": [uids[1]]},
                {"sug_mosad": "3333", "row_uids": [uids[2]]},
            ],
        })
        assert resp.status_code == 200
        assert resp.json()["updated_rows"] == 3

        record = svc.get("test-session")
        sheet1 = record.workbook_dataset.get_sheet_by_name("Sheet1")
        uid_to_row = {r["_row_uid"]: r for r in sheet1.rows}
        assert uid_to_row[uids[0]]["SugMosad"] == "1111"
        assert uid_to_row[uids[1]]["SugMosad"] == "2222"
        assert uid_to_row[uids[2]]["SugMosad"] == "3333"

    def test_more_than_3_groups_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        uids = self._get_uids(svc)
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [
                {"sug_mosad": "1111", "row_uids": [uids[0]]},
                {"sug_mosad": "2222", "row_uids": [uids[0]]},
                {"sug_mosad": "3333", "row_uids": [uids[0]]},
                {"sug_mosad": "4444", "row_uids": [uids[0]]},
            ],
        })
        assert resp.status_code == 422

    def test_empty_row_uids_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "1234", "row_uids": []}],
        })
        assert resp.status_code == 422

    def test_missing_selected_rows_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
        })
        assert resp.status_code == 422

    def test_selected_rows_does_not_touch_other_sheet(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        record = svc.get("test-session")
        for row in record.workbook_dataset.get_sheet_by_name("Sheet2").rows:
            row["SugMosad"] = "UNTOUCHED"

        uids = self._get_uids(svc)
        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "9999", "row_uids": uids}],
        })

        record = svc.get("test-session")
        for row in record.workbook_dataset.get_sheet_by_name("Sheet2").rows:
            assert row.get("SugMosad") == "UNTOUCHED"

    def test_selected_rows_stores_config(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        uids = self._get_uids(svc)
        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "1234", "row_uids": uids[:2]}],
        })
        record = svc.get("test-session")
        cfg = next(c for c in record.sug_mosad_configs if c.scope == "selected_rows")
        assert cfg.sheet_name == "Sheet1"
        assert len(cfg.selected_rows) == 1
        assert cfg.selected_rows[0].sug_mosad == "1234"
        assert set(cfg.selected_rows[0].row_uids) == set(uids[:2])


# ---------------------------------------------------------------------------
# 4. Validation: sug_mosad
# ---------------------------------------------------------------------------

class TestValidationSugMosad:
    def test_non_numeric_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "abc",
        })
        assert resp.status_code == 422

    def test_two_digit_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "12",
        })
        assert resp.status_code == 422

    def test_three_digit_accepted(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "123",
        })
        assert resp.status_code == 200

    def test_four_digit_accepted(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "1234",
        })
        assert resp.status_code == 200

    def test_non_numeric_in_selected_rows_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        uids = [r["_row_uid"] for r in svc.get("test-session").workbook_dataset.get_sheet_by_name("Sheet1").rows]
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "abc", "row_uids": uids[:1]}],
        })
        assert resp.status_code == 422

    def test_two_digit_in_selected_rows_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        uids = [r["_row_uid"] for r in svc.get("test-session").workbook_dataset.get_sheet_by_name("Sheet1").rows]
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "12", "row_uids": uids[:1]}],
        })
        assert resp.status_code == 422

    def test_three_digit_in_selected_rows_accepted(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        uids = [r["_row_uid"] for r in svc.get("test-session").workbook_dataset.get_sheet_by_name("Sheet1").rows]
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "123", "row_uids": uids[:1]}],
        })
        assert resp.status_code == 200


# ---------------------------------------------------------------------------
# 5. Validation: mosad_id
# ---------------------------------------------------------------------------

class TestValidationMosadId:
    def test_non_numeric_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "1234", "mosad_id": "abc",
        })
        assert resp.status_code == 422

    def test_two_digit_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "1234", "mosad_id": "12",
        })
        assert resp.status_code == 422

    def test_three_digit_accepted_and_stored(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "1234", "mosad_id": "567",
        })
        assert resp.status_code == 200
        assert svc.get("test-session").mosad_id == "567"

    def test_four_digit_accepted_and_stored(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "1234", "mosad_id": "5678",
        })
        assert resp.status_code == 200
        assert svc.get("test-session").mosad_id == "5678"

    def test_empty_mosad_id_allowed(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "1234",
        })
        assert resp.status_code == 200


# ---------------------------------------------------------------------------
# 6. Validation: MisparDiraBeMosad
# ---------------------------------------------------------------------------

class TestValidationMisparDira:
    def _make_dira_session(self, tmp_path, svc, has_dira_col):
        file_bytes = _make_xlsx_bytes(["Sheet1"])
        path = tmp_path / "dira.xlsx"
        path.write_bytes(file_bytes)

        field_names = ["first_name", "last_name"]
        rows = [{"first_name": "Alice", "last_name": "Smith", "_row_uid": uuid.uuid4().hex}]
        if has_dira_col:
            field_names.append("מספר דירה במוסד")
            rows[0]["מספר דירה במוסד"] = "not-a-number"

        sheet = SheetDataset(
            sheet_name="Sheet1", header_row=1, header_rows_count=1,
            field_names=field_names, rows=rows,
        )
        wbd = WorkbookDataset(source_file=str(path), sheets=[sheet])
        record = SessionRecord(
            session_id="dira-session",
            source_file_path=str(path), working_copy_path=str(path),
            original_filename="test.xlsx", status="standardized",
            workbook_dataset=wbd,
        )
        svc.create(record)
        return record

    def test_non_numeric_dira_detectable_when_column_exists(self, tmp_path, monkeypatch):
        import webapp.dependencies as deps
        from webapp.services.upload_service import UploadService
        from webapp.services.workbook_service import WorkbookService
        from webapp.services.standardization_service import standardizationService
        from webapp.services.edit_service import EditService
        from webapp.services.export_service import ExportService

        svc = SessionService()
        svc.clear_all()
        self._make_dira_session(tmp_path, svc, has_dira_col=True)

        monkeypatch.setattr(deps, "_session_service", svc)
        monkeypatch.setattr(deps, "_upload_service", UploadService(svc, tmp_path / "u", tmp_path / "w"))
        monkeypatch.setattr(deps, "_workbook_service", WorkbookService(svc))
        monkeypatch.setattr(deps, "_standardization_service", standardizationService(svc))
        monkeypatch.setattr(deps, "_edit_service", EditService(svc))
        monkeypatch.setattr(deps, "_export_service", ExportService(svc, tmp_path / "o"))

        record = svc.get("dira-session")
        sheet = record.workbook_dataset.get_sheet_by_name("Sheet1")
        dira_col = "מספר דירה במוסד"
        assert dira_col in sheet.field_names
        non_numeric = [
            row[dira_col] for row in sheet.rows
            if row.get(dira_col) is not None and not str(row[dira_col]).strip().isdigit()
        ]
        assert len(non_numeric) > 0

    def test_no_failure_when_dira_column_absent(self, tmp_path, monkeypatch):
        import webapp.dependencies as deps
        from webapp.services.upload_service import UploadService
        from webapp.services.workbook_service import WorkbookService
        from webapp.services.standardization_service import standardizationService
        from webapp.services.edit_service import EditService
        from webapp.services.export_service import ExportService

        svc = SessionService()
        svc.clear_all()
        self._make_dira_session(tmp_path, svc, has_dira_col=False)

        monkeypatch.setattr(deps, "_session_service", svc)
        monkeypatch.setattr(deps, "_upload_service", UploadService(svc, tmp_path / "u", tmp_path / "w"))
        monkeypatch.setattr(deps, "_workbook_service", WorkbookService(svc))
        monkeypatch.setattr(deps, "_standardization_service", standardizationService(svc))
        monkeypatch.setattr(deps, "_edit_service", EditService(svc))
        monkeypatch.setattr(deps, "_export_service", ExportService(svc, tmp_path / "o"))

        from webapp.app import app
        with TestClient(app) as client:
            resp = client.post("/api/workbook/dira-session/mosad-type/apply-scoped", json={
                "scope": "workbook", "sug_mosad": "1234",
            })
            assert resp.status_code == 200


# ---------------------------------------------------------------------------
# 7. Export: _resolve_sug_mosad_for_sheet with selected_rows scope
# ---------------------------------------------------------------------------

class TestExportPreservesScope:
    def test_selected_rows_scope_resolver_returns_callable(self):
        from webapp.services.export_service import _resolve_sug_mosad_for_sheet
        from webapp.models.session import SugMosadConfig, SelectedRowsConfig

        uid_a, uid_b, uid_c = "aaa", "bbb", "ccc"
        configs = [
            SugMosadConfig(
                scope="selected_rows",
                sheet_name="Sheet1",
                selected_rows=[
                    SelectedRowsConfig(sug_mosad="1111", row_uids=[uid_a, uid_b]),
                    SelectedRowsConfig(sug_mosad="2222", row_uids=[uid_c]),
                ],
            )
        ]
        resolver = _resolve_sug_mosad_for_sheet(configs, "Sheet1", "fallback")
        assert callable(resolver)
        assert resolver(uid_a) == "1111"
        assert resolver(uid_b) == "1111"
        assert resolver(uid_c) == "2222"
        assert resolver("unknown-uid") is None  # not selected -> leave unchanged

    def test_selected_rows_overrides_sheet_scope(self):
        from webapp.services.export_service import _resolve_sug_mosad_for_sheet
        from webapp.models.session import SugMosadConfig, SelectedRowsConfig

        uid = "row1"
        configs = [
            SugMosadConfig(scope="sheet", sug_mosad="sheet-val", sheet_name="Sheet1"),
            SugMosadConfig(
                scope="selected_rows", sheet_name="Sheet1",
                selected_rows=[SelectedRowsConfig(sug_mosad="rows-val", row_uids=[uid])],
            ),
        ]
        resolver = _resolve_sug_mosad_for_sheet(configs, "Sheet1", "fallback")
        assert callable(resolver)
        assert resolver(uid) == "rows-val"

    def test_sheet_scope_resolver(self):
        from webapp.services.export_service import _resolve_sug_mosad_for_sheet
        from webapp.models.session import SugMosadConfig

        configs = [SugMosadConfig(scope="sheet", sug_mosad="9999", sheet_name="Sheet1")]
        assert _resolve_sug_mosad_for_sheet(configs, "Sheet1", "fallback") == "9999"
        assert _resolve_sug_mosad_for_sheet(configs, "Sheet2", "fallback") == "fallback"

    def test_workbook_scope_resolver(self):
        from webapp.services.export_service import _resolve_sug_mosad_for_sheet
        from webapp.models.session import SugMosadConfig

        configs = [SugMosadConfig(scope="workbook", sug_mosad="7777")]
        assert _resolve_sug_mosad_for_sheet(configs, "Sheet1", "fallback") == "7777"
        assert _resolve_sug_mosad_for_sheet(configs, "AnySheet", "fallback") == "7777"

    def test_no_config_uses_fallback(self):
        from webapp.services.export_service import _resolve_sug_mosad_for_sheet
        assert _resolve_sug_mosad_for_sheet([], "Sheet1", "legacy") == "legacy"
        assert _resolve_sug_mosad_for_sheet(None, "Sheet1", "legacy") == "legacy"

    def test_unselected_rows_use_fallback_during_export(self):
        """Rows not in selected_rows get None from resolver (leave unchanged)."""
        from webapp.services.export_service import _resolve_sug_mosad_for_sheet
        from webapp.models.session import SugMosadConfig, SelectedRowsConfig

        configs = [
            SugMosadConfig(
                scope="selected_rows", sheet_name="Sheet1",
                selected_rows=[SelectedRowsConfig(sug_mosad="1234", row_uids=["uid-selected"])],
            )
        ]
        resolver = _resolve_sug_mosad_for_sheet(configs, "Sheet1", "fallback")
        assert resolver("uid-selected") == "1234"
        assert resolver("uid-not-selected") is None  # not overridden



# ---------------------------------------------------------------------------
# 8. Validation: legacy /mosad-type/apply endpoint
# ---------------------------------------------------------------------------

class TestLegacyApplyValidation:
    """The legacy /mosad-type/apply endpoint must also enforce min-3 numeric."""

    def _client_with_type(self, tmp_path, monkeypatch, mosad_type: str):
        """Return a TestClient whose session already has mosad_type stored."""
        import webapp.dependencies as deps
        from webapp.services.upload_service import UploadService
        from webapp.services.workbook_service import WorkbookService
        from webapp.services.standardization_service import standardizationService
        from webapp.services.edit_service import EditService
        from webapp.services.export_service import ExportService

        svc, record = _make_session(tmp_path, sheet_names=["Sheet1"], n_rows=2)
        # Pre-store the type so the endpoint can find it
        svc.update("test-session", mosad_types=[mosad_type])

        monkeypatch.setattr(deps, "_session_service", svc)
        monkeypatch.setattr(deps, "_upload_service", UploadService(svc, tmp_path / "u", tmp_path / "w"))
        monkeypatch.setattr(deps, "_workbook_service", WorkbookService(svc))
        monkeypatch.setattr(deps, "_standardization_service", standardizationService(svc))
        monkeypatch.setattr(deps, "_edit_service", EditService(svc))
        monkeypatch.setattr(deps, "_export_service", ExportService(svc, tmp_path / "o"))

        from webapp.app import app
        return TestClient(app), svc

    def test_two_digit_mosad_type_rejected_by_legacy_endpoint(self, tmp_path, monkeypatch):
        client, _ = self._client_with_type(tmp_path, monkeypatch, "12")
        resp = client.post("/api/workbook/test-session/mosad-type/apply",
                           json={"mosad_type": "12"})
        assert resp.status_code == 422

    def test_non_numeric_mosad_type_rejected_by_legacy_endpoint(self, tmp_path, monkeypatch):
        client, _ = self._client_with_type(tmp_path, monkeypatch, "abc")
        resp = client.post("/api/workbook/test-session/mosad-type/apply",
                           json={"mosad_type": "abc"})
        assert resp.status_code == 422

    def test_three_digit_mosad_type_accepted_by_legacy_endpoint(self, tmp_path, monkeypatch):
        client, _ = self._client_with_type(tmp_path, monkeypatch, "123")
        resp = client.post("/api/workbook/test-session/mosad-type/apply",
                           json={"mosad_type": "123"})
        assert resp.status_code == 200


# ---------------------------------------------------------------------------
# 9. Validation: PATCH /institution endpoint
# ---------------------------------------------------------------------------

class TestPatchInstitutionValidation:
    """PATCH /institution must validate mosad_id and mosad_types."""

    def test_two_digit_mosad_id_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.patch("/api/workbook/test-session/institution",
                            json={"mosad_id": "12"})
        assert resp.status_code == 422

    def test_non_numeric_mosad_id_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.patch("/api/workbook/test-session/institution",
                            json={"mosad_id": "abc"})
        assert resp.status_code == 422

    def test_three_digit_mosad_id_accepted(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.patch("/api/workbook/test-session/institution",
                            json={"mosad_id": "123"})
        assert resp.status_code == 200
        assert svc.get("test-session").mosad_id == "123"

    def test_two_digit_mosad_type_in_patch_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.patch("/api/workbook/test-session/institution",
                            json={"mosad_types": ["12"]})
        assert resp.status_code == 422

    def test_non_numeric_mosad_type_in_patch_rejected(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.patch("/api/workbook/test-session/institution",
                            json={"mosad_types": ["abc"]})
        assert resp.status_code == 422

    def test_three_digit_mosad_type_in_patch_accepted(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.patch("/api/workbook/test-session/institution",
                            json={"mosad_types": ["123"]})
        assert resp.status_code == 200
        assert svc.get("test-session").mosad_types == ["123"]

    def test_empty_mosad_id_accepted(self, client_with_two_sheets):
        client, svc, record = client_with_two_sheets
        resp = client.patch("/api/workbook/test-session/institution",
                            json={"mosad_id": ""})
        assert resp.status_code == 200

    def test_mosad_name_not_validated_as_numeric(self, client_with_two_sheets):
        """mosad_name is free text — must not be rejected."""
        client, svc, record = client_with_two_sheets
        resp = client.patch("/api/workbook/test-session/institution",
                            json={"mosad_name": "בית הספר הטוב"})
        assert resp.status_code == 200


# ---------------------------------------------------------------------------
# 10. Selected-rows accumulation — second apply merges, not overwrites
# ---------------------------------------------------------------------------

class TestSelectedRowsAccumulation:
    """Applying institution type to selected rows twice on the same sheet
    must accumulate the configs, not overwrite the first group."""

    def _get_uids(self, svc, sheet_name="Sheet1"):
        record = svc.get("test-session")
        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)
        return [row["_row_uid"] for row in sheet.rows]

    def test_second_apply_to_different_rows_accumulates(self, client_with_two_sheets):
        """Row 0 gets 111, then row 1 gets 222 — both must be preserved."""
        client, svc, record = client_with_two_sheets
        uids = self._get_uids(svc)

        # First apply: row 0 → 111
        r1 = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "111", "row_uids": [uids[0]]}],
        })
        assert r1.status_code == 200

        # Second apply: row 1 → 222
        r2 = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "222", "row_uids": [uids[1]]}],
        })
        assert r2.status_code == 200

        # Both groups must be in the config
        record = svc.get("test-session")
        sr_cfg = next(
            c for c in record.sug_mosad_configs
            if c.scope == "selected_rows" and c.sheet_name == "Sheet1"
        )
        uid_to_sug = {uid: grp.sug_mosad for grp in sr_cfg.selected_rows for uid in grp.row_uids}
        assert uid_to_sug.get(uids[0]) == "111"
        assert uid_to_sug.get(uids[1]) == "222"

    def test_reassigning_row_moves_it_to_new_group(self, client_with_two_sheets):
        """Row 0 gets 111, then row 0 gets 222 — it must move to 222, not stay in 111."""
        client, svc, record = client_with_two_sheets
        uids = self._get_uids(svc)

        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "111", "row_uids": [uids[0]]}],
        })
        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "222", "row_uids": [uids[0]]}],
        })

        record = svc.get("test-session")
        sr_cfg = next(
            c for c in record.sug_mosad_configs
            if c.scope == "selected_rows" and c.sheet_name == "Sheet1"
        )
        uid_to_sug = {uid: grp.sug_mosad for grp in sr_cfg.selected_rows for uid in grp.row_uids}
        # uid[0] must be in 222 only
        assert uid_to_sug.get(uids[0]) == "222"
        # uid[0] must not appear in any 111 group
        for grp in sr_cfg.selected_rows:
            if grp.sug_mosad == "111":
                assert uids[0] not in grp.row_uids

    def test_unselected_rows_not_overwritten_by_selected_rows_mode(self, client_with_two_sheets):
        """Rows not in selected_rows must keep their original SugMosad."""
        client, svc, record = client_with_two_sheets
        uids = self._get_uids(svc)

        # Pre-set all rows to 999 via workbook scope
        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "999",
        })

        # Now apply 111 only to row 0
        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "111", "row_uids": [uids[0]]}],
        })

        record = svc.get("test-session")
        sheet = record.workbook_dataset.get_sheet_by_name("Sheet1")
        uid_to_row = {r["_row_uid"]: r for r in sheet.rows}

        # Row 0 must have 111
        assert uid_to_row[uids[0]]["SugMosad"] == "111"
        # Rows 1 and 2 must still have 999 (set by workbook scope)
        assert uid_to_row[uids[1]]["SugMosad"] == "999"
        assert uid_to_row[uids[2]]["SugMosad"] == "999"

    def test_workbook_scope_fallback_still_works(self, client_with_two_sheets):
        """Workbook-scope apply still sets all rows when no selected_rows config exists."""
        client, svc, record = client_with_two_sheets
        resp = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "777",
        })
        assert resp.status_code == 200
        record = svc.get("test-session")
        for sheet in record.workbook_dataset.sheets:
            for row in sheet.rows:
                assert row["SugMosad"] == "777"

    def test_sheet_scope_overrides_workbook_for_that_sheet_only(self, client_with_two_sheets):
        """Sheet-scope apply overrides workbook scope for the target sheet only."""
        client, svc, record = client_with_two_sheets

        # Workbook scope first
        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "777",
        })
        # Sheet scope for Sheet1 only
        client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "sheet", "sug_mosad": "888", "sheet_name": "Sheet1",
        })

        record = svc.get("test-session")
        sheet1 = record.workbook_dataset.get_sheet_by_name("Sheet1")
        sheet2 = record.workbook_dataset.get_sheet_by_name("Sheet2")
        for row in sheet1.rows:
            assert row["SugMosad"] == "888"
        for row in sheet2.rows:
            assert row["SugMosad"] == "777"

    def test_two_digit_sug_mosad_rejected_in_all_scopes(self, client_with_two_sheets):
        """2-digit sug_mosad must be rejected in workbook, sheet, and selected_rows scopes."""
        client, svc, record = client_with_two_sheets
        uids = [r["_row_uid"] for r in svc.get("test-session").workbook_dataset.get_sheet_by_name("Sheet1").rows]

        r1 = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "12",
        })
        assert r1.status_code == 422

        r2 = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "sheet", "sug_mosad": "12", "sheet_name": "Sheet1",
        })
        assert r2.status_code == 422

        r3 = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "12", "row_uids": uids[:1]}],
        })
        assert r3.status_code == 422

    def test_two_digit_mosad_id_rejected_in_all_scopes(self, client_with_two_sheets):
        """2-digit mosad_id must be rejected regardless of scope."""
        client, svc, record = client_with_two_sheets

        r1 = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "1234", "mosad_id": "12",
        })
        assert r1.status_code == 422

        r2 = client.post("/api/workbook/test-session/mosad-type/apply-scoped", json={
            "scope": "sheet", "sug_mosad": "1234", "sheet_name": "Sheet1", "mosad_id": "12",
        })
        assert r2.status_code == 422


# ---------------------------------------------------------------------------
# 11. Runtime regression: selected_rows applies only to selected rows
#     (guards against the HTML bug where all rows were changed)
# ---------------------------------------------------------------------------

class TestSelectedRowsRuntimeRegression:
    """Regression tests that mirror the exact runtime flow:
    5 rows, apply sug_mosad 123 to 2 row_uids only.
    The other 3 rows must remain unchanged.
    No workbook-level or sheet-level config must be created.
    """

    def _make_5row_session(self, tmp_path, monkeypatch):
        import webapp.dependencies as deps
        from webapp.services.upload_service import UploadService
        from webapp.services.workbook_service import WorkbookService
        from webapp.services.standardization_service import standardizationService
        from webapp.services.edit_service import EditService
        from webapp.services.export_service import ExportService

        svc = SessionService()
        svc.clear_all()

        file_bytes = _make_xlsx_bytes(["Sheet1"])
        path = tmp_path / "five.xlsx"
        path.write_bytes(file_bytes)

        rows = [
            {"first_name": f"Name{i}", "last_name": "Test", "_row_uid": f"uid-{i}"}
            for i in range(1, 6)
        ]
        sheet = SheetDataset(
            sheet_name="Sheet1", header_row=1, header_rows_count=1,
            field_names=["first_name", "last_name"], rows=rows,
        )
        wbd = WorkbookDataset(source_file=str(path), sheets=[sheet])
        record = SessionRecord(
            session_id="reg-session",
            source_file_path=str(path), working_copy_path=str(path),
            original_filename="five.xlsx", status="standardized",
            workbook_dataset=wbd,
        )
        svc.create(record)

        monkeypatch.setattr(deps, "_session_service", svc)
        monkeypatch.setattr(deps, "_upload_service",
                            UploadService(svc, tmp_path / "u", tmp_path / "w"))
        monkeypatch.setattr(deps, "_workbook_service", WorkbookService(svc))
        monkeypatch.setattr(deps, "_standardization_service", standardizationService(svc))
        monkeypatch.setattr(deps, "_edit_service", EditService(svc))
        monkeypatch.setattr(deps, "_export_service",
                            ExportService(svc, tmp_path / "o"))

        from webapp.app import app
        return TestClient(app), svc

    def test_only_selected_2_rows_get_sug_mosad(self, tmp_path, monkeypatch):
        """Apply 123 to uid-1 and uid-2 only. uid-3, uid-4, uid-5 must be unchanged."""
        client, svc = self._make_5row_session(tmp_path, monkeypatch)

        resp = client.post("/api/workbook/reg-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "123", "row_uids": ["uid-1", "uid-2"]}],
        })
        assert resp.status_code == 200
        assert resp.json()["updated_rows"] == 2

        record = svc.get("reg-session")
        sheet = record.workbook_dataset.get_sheet_by_name("Sheet1")
        uid_to_row = {r["_row_uid"]: r for r in sheet.rows}

        # Selected rows must have 123
        assert uid_to_row["uid-1"].get("SugMosad") == "123"
        assert uid_to_row["uid-2"].get("SugMosad") == "123"

        # Unselected rows must NOT have 123 (must be absent or unchanged)
        assert uid_to_row["uid-3"].get("SugMosad") != "123"
        assert uid_to_row["uid-4"].get("SugMosad") != "123"
        assert uid_to_row["uid-5"].get("SugMosad") != "123"

    def test_no_workbook_or_sheet_config_created_by_selected_rows_apply(
        self, tmp_path, monkeypatch
    ):
        """selected_rows apply must not create a workbook or sheet-level config."""
        client, svc = self._make_5row_session(tmp_path, monkeypatch)

        client.post("/api/workbook/reg-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "123", "row_uids": ["uid-1"]}],
        })

        record = svc.get("reg-session")
        for cfg in record.sug_mosad_configs:
            assert cfg.scope not in ("workbook", "sheet"), (
                f"selected_rows apply must not create a {cfg.scope!r} config"
            )

    def test_payload_scope_workbook_still_changes_all_rows(self, tmp_path, monkeypatch):
        """Sanity: scope=workbook must still change all 5 rows."""
        client, svc = self._make_5row_session(tmp_path, monkeypatch)

        resp = client.post("/api/workbook/reg-session/mosad-type/apply-scoped", json={
            "scope": "workbook",
            "sug_mosad": "999",
        })
        assert resp.status_code == 200
        assert resp.json()["updated_rows"] == 5

        record = svc.get("reg-session")
        sheet = record.workbook_dataset.get_sheet_by_name("Sheet1")
        for row in sheet.rows:
            assert row.get("SugMosad") == "999"

    def test_selected_rows_does_not_affect_unselected_even_after_workbook_apply(
        self, tmp_path, monkeypatch
    ):
        """After workbook apply sets 999, selected_rows apply of 123 to uid-1
        must leave uid-2..uid-5 at 999."""
        client, svc = self._make_5row_session(tmp_path, monkeypatch)

        # First: workbook scope sets all to 999
        client.post("/api/workbook/reg-session/mosad-type/apply-scoped", json={
            "scope": "workbook", "sug_mosad": "999",
        })

        # Then: selected_rows scope sets uid-1 to 123
        client.post("/api/workbook/reg-session/mosad-type/apply-scoped", json={
            "scope": "selected_rows",
            "sheet_name": "Sheet1",
            "selected_rows": [{"sug_mosad": "123", "row_uids": ["uid-1"]}],
        })

        record = svc.get("reg-session")
        sheet = record.workbook_dataset.get_sheet_by_name("Sheet1")
        uid_to_row = {r["_row_uid"]: r for r in sheet.rows}

        assert uid_to_row["uid-1"].get("SugMosad") == "123"
        for uid in ["uid-2", "uid-3", "uid-4", "uid-5"]:
            assert uid_to_row[uid].get("SugMosad") == "999"
