"""Unit tests for WorkbookService."""

import pytest
from fastapi import HTTPException

from src.excel_standardization.data_types import SheetDataset, WorkbookDataset
from webapp.models.session import SessionRecord
from webapp.services.session_service import SessionService
from webapp.services.workbook_service import WorkbookService


def make_workbook_dataset():
    sheet1 = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "last_name"],
        rows=[
            {"first_name": "Alice", "last_name": "Smith"},
            {"first_name": "Bob", "last_name": "Jones"},
        ],
    )
    sheet2 = SheetDataset(
        sheet_name="Sheet2",
        header_row=1,
        header_rows_count=1,
        field_names=["gender"],
        rows=[{"gender": "F"}],
    )
    return WorkbookDataset(source_file="test.xlsx", sheets=[sheet1, sheet2])


@pytest.fixture(autouse=True)
def clear_registry():
    svc = SessionService()
    svc.clear_all()
    yield
    svc.clear_all()


@pytest.fixture
def session_with_workbook():
    svc = SessionService()
    wb = make_workbook_dataset()
    record = SessionRecord(
        session_id="wb-session",
        source_file_path="uploads/wb-session.xlsx",
        working_copy_path="work/wb-session.xlsx",
        original_filename="test.xlsx",
        status="uploaded",
        workbook_dataset=wb,
    )
    svc.create(record)
    return svc, WorkbookService(svc)


def test_get_summary_returns_correct_sheet_names(session_with_workbook):
    _, wb_svc = session_with_workbook
    summary = wb_svc.get_summary("wb-session")
    assert summary.session_id == "wb-session"
    assert len(summary.sheets) == 2
    names = [s.sheet_name for s in summary.sheets]
    assert "Sheet1" in names
    assert "Sheet2" in names


def test_get_summary_returns_correct_row_counts(session_with_workbook):
    _, wb_svc = session_with_workbook
    summary = wb_svc.get_summary("wb-session")
    sheet1_summary = next(s for s in summary.sheets if s.sheet_name == "Sheet1")
    assert sheet1_summary.row_count == 2
    sheet2_summary = next(s for s in summary.sheets if s.sheet_name == "Sheet2")
    assert sheet2_summary.row_count == 1


def test_get_summary_returns_correct_field_names(session_with_workbook):
    _, wb_svc = session_with_workbook
    summary = wb_svc.get_summary("wb-session")
    sheet1_summary = next(s for s in summary.sheets if s.sheet_name == "Sheet1")
    assert "first_name" in sheet1_summary.field_names
    assert "last_name" in sheet1_summary.field_names


def test_get_sheet_data_returns_rows_for_valid_sheet(session_with_workbook):
    _, wb_svc = session_with_workbook
    response = wb_svc.get_sheet_data("wb-session", "Sheet1")
    assert response.sheet_name == "Sheet1"
    assert len(response.rows) == 2
    assert response.rows[0]["first_name"] == "Alice"


def test_get_sheet_data_shows_generated_passport_corrected_without_source_passport_column():
    svc = SessionService()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["id_number", "passport_corrected"],
        rows=[
            {
                "id_number": "ABC123",
                "id_number_corrected": "",
                "passport_corrected": "ABC123",
                "identifier_status": "moved",
            }
        ],
    )
    record = SessionRecord(
        session_id="passport-ui-session",
        source_file_path="uploads/passport-ui-session.xlsx",
        working_copy_path="work/passport-ui-session.xlsx",
        original_filename="test.xlsx",
        status="standardized",
        workbook_dataset=WorkbookDataset(source_file="test.xlsx", sheets=[sheet]),
    )
    svc.create(record)

    response = WorkbookService(svc).get_sheet_data("passport-ui-session", "Sheet1")

    id_index = response.field_names.index("id_number")
    assert response.field_names[id_index + 1] == "id_number_corrected"
    assert response.field_names[id_index + 2] == "passport_corrected"
    assert response.field_names[id_index + 3] == "identifier_status"
    assert "passport_corrected" in response.field_names
    assert response.rows[0]["passport_corrected"] == "ABC123"
    assert "passport" not in response.field_names
    assert "passport" not in response.rows[0]


def test_get_sheet_data_keeps_numeric_invalid_length_id_out_of_passport():
    svc = SessionService()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "id_number"],
        rows=[
            {
                "first_name": "Person",
                "id_number": "1234567890",
                "id_number_corrected": "1234567890",
                "identifier_status": "ת.ז. לא תקינה",
            }
        ],
    )
    record = SessionRecord(
        session_id="identifier-ui-session",
        source_file_path="uploads/identifier-ui-session.xlsx",
        working_copy_path="work/identifier-ui-session.xlsx",
        original_filename="test.xlsx",
        status="standardized",
        workbook_dataset=WorkbookDataset(source_file="test.xlsx", sheets=[sheet]),
    )
    svc.create(record)

    response = WorkbookService(svc).get_sheet_data("identifier-ui-session", "Sheet1")

    id_index = response.field_names.index("id_number")
    assert response.field_names[id_index + 1] == "id_number_corrected"
    assert response.field_names[id_index + 2] == "identifier_status"
    assert "passport_corrected" not in response.field_names
    assert response.rows[0]["id_number"] == "1234567890"
    assert response.rows[0]["id_number_corrected"] == "1234567890"


def test_get_sheet_data_orders_name_corrected_fields_after_source_fields():
    svc = SessionService()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "last_name", "father_name"],
        rows=[
            {
                "first_name": "Alice Smith",
                "first_name_corrected": "Alice",
                "last_name": "Smith",
                "last_name_corrected": "Smith",
                "father_name": "Robert Smith",
                "father_name_corrected": "Robert",
            }
        ],
    )
    record = SessionRecord(
        session_id="name-ui-session",
        source_file_path="uploads/name-ui-session.xlsx",
        working_copy_path="work/name-ui-session.xlsx",
        original_filename="test.xlsx",
        status="standardized",
        workbook_dataset=WorkbookDataset(source_file="test.xlsx", sheets=[sheet]),
    )
    svc.create(record)

    response = WorkbookService(svc).get_sheet_data("name-ui-session", "Sheet1")

    for source, corrected in [
        ("first_name", "first_name_corrected"),
        ("last_name", "last_name_corrected"),
        ("father_name", "father_name_corrected"),
    ]:
        source_index = response.field_names.index(source)
        assert response.field_names[source_index + 1] == corrected
    assert response.rows[0]["first_name_corrected"] == "Alice"
    assert response.rows[0]["last_name_corrected"] == "Smith"
    assert response.rows[0]["father_name_corrected"] == "Robert"


def test_get_sheet_data_orders_gender_corrected_and_status_after_source_field():
    svc = SessionService()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "gender"],
        rows=[
            {
                "first_name": "Person",
                "gender": "8",
                "gender_corrected": "",
                "gender_status": "קוד מין לא תקין - חייב להיות 1 (זכר) או 2 (נקבה)",
            }
        ],
    )
    record = SessionRecord(
        session_id="gender-ui-session",
        source_file_path="uploads/gender-ui-session.xlsx",
        working_copy_path="work/gender-ui-session.xlsx",
        original_filename="test.xlsx",
        status="standardized",
        workbook_dataset=WorkbookDataset(source_file="test.xlsx", sheets=[sheet]),
    )
    svc.create(record)

    response = WorkbookService(svc).get_sheet_data("gender-ui-session", "Sheet1")

    gender_index = response.field_names.index("gender")
    assert response.field_names[gender_index + 1] == "gender_corrected"
    assert response.field_names[gender_index + 2] == "gender_status"
    assert response.rows[0]["first_name"] == "Person"
    assert response.rows[0]["gender"] == "8"
    assert response.rows[0]["gender_corrected"] == ""
    assert response.rows[0]["gender_status"] == "קוד מין לא תקין - חייב להיות 1 (זכר) או 2 (נקבה)"


def test_get_sheet_data_preserves_original_order_with_corrected_fields_and_status_anchors():
    svc = SessionService()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "id_number", "passport", "gender", "last_name"],
        rows=[
            {
                "first_name": "Original First",
                "first_name_corrected": "Corrected First",
                "id_number": "123",
                "id_number_corrected": "000000123",
                "passport": "P1",
                "passport_corrected": "P1",
                "identifier_status": "identifier status",
                "gender": "F",
                "gender_corrected": 2,
                "gender_status": "gender status",
                "last_name": "Original Last",
                "last_name_corrected": "Corrected Last",
            }
        ],
    )
    record = SessionRecord(
        session_id="ordering-contract-session",
        source_file_path="uploads/ordering-contract-session.xlsx",
        working_copy_path="work/ordering-contract-session.xlsx",
        original_filename="test.xlsx",
        status="standardized",
        workbook_dataset=WorkbookDataset(source_file="test.xlsx", sheets=[sheet]),
    )
    svc.create(record)

    response = WorkbookService(svc).get_sheet_data("ordering-contract-session", "Sheet1")
    fields = [field for field in response.field_names if field != "_serial"]

    assert fields == [
        "first_name",
        "first_name_corrected",
        "id_number",
        "id_number_corrected",
        "passport",
        "passport_corrected",
        "identifier_status",
        "gender",
        "gender_corrected",
        "gender_status",
        "last_name",
        "last_name_corrected",
    ]


def test_get_sheet_data_date_groups_remain_block_based():
    svc = SessionService()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "birth_year", "birth_month", "birth_day", "gender"],
        rows=[
            {
                "first_name": "Person",
                "first_name_corrected": "Person",
                "birth_year": "1980",
                "birth_month": "05",
                "birth_day": "17",
                "birth_year_corrected": 1980,
                "birth_month_corrected": 5,
                "birth_day_corrected": 17,
                "birth_date_status": "",
                "gender": "M",
                "gender_corrected": 1,
                "gender_status": "",
            }
        ],
    )
    record = SessionRecord(
        session_id="date-ordering-session",
        source_file_path="uploads/date-ordering-session.xlsx",
        working_copy_path="work/date-ordering-session.xlsx",
        original_filename="test.xlsx",
        status="standardized",
        workbook_dataset=WorkbookDataset(source_file="test.xlsx", sheets=[sheet]),
    )
    svc.create(record)

    response = WorkbookService(svc).get_sheet_data("date-ordering-session", "Sheet1")
    fields = [field for field in response.field_names if field != "_serial"]

    assert fields == [
        "first_name",
        "first_name_corrected",
        "birth_year",
        "birth_month",
        "birth_day",
        "birth_year_corrected",
        "birth_month_corrected",
        "birth_day_corrected",
        "birth_date_status",
        "gender",
        "gender_corrected",
        "gender_status",
    ]


def test_get_sheet_data_helper_row_filtering_keeps_ordering_contract():
    svc = SessionService()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["id_number", "gender"],
        rows=[
            {"id_number": "1", "gender": "2"},
            {
                "id_number": "ABC123",
                "id_number_corrected": "",
                "passport_corrected": "ABC123",
                "identifier_status": "moved",
                "gender": "F",
                "gender_corrected": 2,
                "gender_status": "",
            },
        ],
    )
    record = SessionRecord(
        session_id="helper-ordering-session",
        source_file_path="uploads/helper-ordering-session.xlsx",
        working_copy_path="work/helper-ordering-session.xlsx",
        original_filename="test.xlsx",
        status="standardized",
        workbook_dataset=WorkbookDataset(source_file="test.xlsx", sheets=[sheet]),
    )
    svc.create(record)

    response = WorkbookService(svc).get_sheet_data("helper-ordering-session", "Sheet1")
    fields = [field for field in response.field_names if field != "_serial"]

    assert len(response.rows) == 1
    assert response.rows[0]["id_number"] == "ABC123"
    assert fields == [
        "id_number",
        "id_number_corrected",
        "passport_corrected",
        "identifier_status",
        "gender",
        "gender_corrected",
        "gender_status",
    ]


def test_get_sheet_data_raises_404_for_unknown_sheet(session_with_workbook):
    _, wb_svc = session_with_workbook
    with pytest.raises(HTTPException) as exc_info:
        wb_svc.get_sheet_data("wb-session", "NonExistentSheet")
    assert exc_info.value.status_code == 404


def test_get_summary_raises_404_for_unknown_session(session_with_workbook):
    _, wb_svc = session_with_workbook
    with pytest.raises(HTTPException) as exc_info:
        wb_svc.get_summary("ghost-session")
    assert exc_info.value.status_code == 404
