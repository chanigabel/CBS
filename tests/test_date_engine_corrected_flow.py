"""Corrected DateEngine flow guarantees for the active Dataset/Web path."""

from datetime import date

from src.excel_standardization.data_types import (
    DateInput,
    DateFieldType,
    DateFormatPattern,
    SheetDataset,
)
from src.excel_standardization.engines.date_engine import DateEngine
from src.excel_standardization.export.export_engine import ExportEngine
from src.excel_standardization.processing.standardization_pipeline import StandardizationPipeline


REFERENCE_DATE = date(2026, 5, 11)


def _pipeline() -> StandardizationPipeline:
    return StandardizationPipeline(
        date_engine=DateEngine(reference_date=REFERENCE_DATE),
        apply_name_standardization_enabled=False,
        apply_gender_standardization_enabled=False,
        apply_date_standardization_enabled=True,
        apply_identifier_standardization_enabled=False,
        reference_date=REFERENCE_DATE,
    )


def _dataset(rows, fields, sheet_name="Dayarim"):
    return SheetDataset(
        sheet_name=sheet_name,
        header_row=1,
        header_rows_count=1,
        field_names=fields,
        rows=rows,
        metadata={},
    )


def test_partial_split_preserves_year_month_and_marks_missing_day():
    row = {"birth_year": "2010", "birth_month": "05", "birth_day": ""}

    _pipeline().apply_date_standardization(row)

    assert row["birth_year_corrected"] == 2010
    assert row["birth_month_corrected"] == 5
    assert row["birth_day_corrected"] == ""
    assert row["birth_date_status"] == "חסר יום"


def test_partial_split_preserves_month_day_and_marks_missing_year():
    row = {"birth_year": "", "birth_month": "05", "birth_day": "12"}

    _pipeline().apply_date_standardization(row)

    assert row["birth_year_corrected"] == ""
    assert row["birth_month_corrected"] == 5
    assert row["birth_day_corrected"] == 12
    assert row["birth_date_status"] == "חסר שנה"


def test_single_invalid_date_uses_same_export_safe_component_policy():
    row = {"birth_date": "31/02/2020"}

    _pipeline().apply_date_standardization(row)

    assert row["birth_year_corrected"] == 2020
    assert row["birth_month_corrected"] == 2
    assert row["birth_day_corrected"] == ""
    assert row["birth_date_status"] == "תאריך לא קיים"


def test_two_digit_year_metadata_is_set_for_numeric_and_separated_paths():
    engine = DateEngine(reference_date=REFERENCE_DATE)

    numeric = engine.parse_date(
        None,
        None,
        None,
        "010224",
        DateFormatPattern.DDMM,
        DateFieldType.BIRTH_DATE,
    )
    separated = engine.parse_date(
        None,
        None,
        None,
        "01/02/24",
        DateFormatPattern.DDMM,
        DateFieldType.BIRTH_DATE,
    )

    assert numeric.year == 2024
    assert numeric.year_was_auto_completed is True
    assert numeric.original_year_value == 24
    assert numeric.reference_year == 2026
    assert separated.year == 2024
    assert separated.year_was_auto_completed is True
    assert separated.original_year_value == 24
    assert separated.reference_year == 2026


def test_reference_year_makes_two_digit_expansion_deterministic():
    engine_2026 = DateEngine(reference_date=date(2026, 1, 1))
    engine_2027 = DateEngine(reference_date=date(2027, 1, 1))

    result_2026 = engine_2026.parse_separated_date_string("01/01/27", DateFormatPattern.DDMM)
    result_2027 = engine_2027.parse_separated_date_string("01/01/27", DateFormatPattern.DDMM)

    assert result_2026.year == 1927
    assert result_2027.year == 2027


def test_majority_correction_includes_single_numeric_and_separated_dates():
    ds = _dataset(
        [
            {"birth_date": "010130"},
            {"birth_date": "010135"},
            {"birth_date": "010124"},
        ],
        ["birth_date"],
    )

    normalized = _pipeline().normalize_dataset(ds)

    years = [row["birth_year_corrected"] for row in normalized.rows]
    assert years == [1930, 1935, 1924]


def test_pipeline_records_processing_date_metadata():
    ds = _dataset([{"birth_date": "01/02/24"}], ["birth_date"])

    normalized = _pipeline().normalize_dataset(ds)

    assert normalized.metadata["processing_date"] == "2026-05-11"
    assert normalized.metadata["processing_year"] == 2026


def test_dataset_export_uses_corrected_dates_only_without_original_fallback():
    row = {
        "birth_year": 1999,
        "birth_month": 12,
        "birth_day": 31,
        "birth_year_corrected": "",
        "birth_month_corrected": "",
        "birth_day_corrected": "",
    }

    mapped = ExportEngine()._map_row_to_export_fields(row, "Dayarim", allow_mosad_fields=True)

    assert mapped["ShnatLida"] == ""
    assert mapped["HodeshLida"] == ""
    assert mapped["YomLida"] == ""


def test_compact_numeric_mmdd_fallbacks_and_invalid_values_blank_components():
    engine = DateEngine(reference_date=REFERENCE_DATE)

    eight_digit = engine.parse_date(
        None, None, None, "12312024", DateFormatPattern.DDMM, DateFieldType.BIRTH_DATE
    )
    six_digit = engine.parse_date(
        None, None, None, "123124", DateFormatPattern.DDMM, DateFieldType.BIRTH_DATE
    )
    invalid = _pipeline().normalize_row({"birth_date": "999999"})

    assert (eight_digit.year, eight_digit.month, eight_digit.day) == (2024, 12, 31)
    assert (six_digit.year, six_digit.month, six_digit.day) == (2024, 12, 31)
    assert invalid["birth_year_corrected"] == ""
    assert invalid["birth_month_corrected"] == ""
    assert invalid["birth_day_corrected"] == ""
    assert invalid["birth_date_status"] != ""


def test_separator_normalization_trailing_text_and_split_zero_recovery():
    pipeline = _pipeline()

    repeated = pipeline.normalize_row({"birth_date": "01//02//2024"})
    trailing = pipeline.normalize_row({"birth_date": "01/02/2024abc"})
    split = pipeline.normalize_row(
        {"birth_year": 0, "birth_month": "", "birth_day": "11.06.1997"}
    )

    assert (
        repeated["birth_year_corrected"],
        repeated["birth_month_corrected"],
        repeated["birth_day_corrected"],
    ) == (2024, 2, 1)
    assert (
        trailing["birth_year_corrected"],
        trailing["birth_month_corrected"],
        trailing["birth_day_corrected"],
    ) == (2024, 2, 1)
    assert trailing["birth_date_status"] != ""
    assert (
        split["birth_year_corrected"],
        split["birth_month_corrected"],
        split["birth_day_corrected"],
    ) == (1997, 6, 11)
    assert split["birth_date_status"] != ""


def test_excel_serial_status_survives_business_warning():
    result = DateEngine(reference_date=REFERENCE_DATE).parse_input(
        DateInput(
            source_kind="single",
            field_type=DateFieldType.BIRTH_DATE,
            raw_value=3000,
            pattern=DateFormatPattern.DDMM,
            reference_date=REFERENCE_DATE,
            source_is_excel_date_serial=True,
        )
    )

    assert "פורק מתאריך סידורי" in result.status_text
    assert "גיל מעל 100" in result.status_text


def test_statuses_do_not_erase_parsed_components_in_pipeline():
    pipeline = _pipeline()

    year_only = pipeline.normalize_row({"birth_date": "2020"})
    calendar_invalid = pipeline.normalize_row({"birth_date": "31022020"})
    future_birth = pipeline.normalize_row({"birth_date": "12/05/2026"})
    before_1906 = pipeline.normalize_row({"birth_date": "01/01/1900"})
    late_entry = pipeline.normalize_row({"entry_date": "01/01/2026"})

    assert (
        year_only["birth_year_corrected"],
        year_only["birth_month_corrected"],
        year_only["birth_day_corrected"],
    ) == (2020, "", "")
    assert year_only["birth_date_status"] != ""

    assert (
        calendar_invalid["birth_year_corrected"],
        calendar_invalid["birth_month_corrected"],
        calendar_invalid["birth_day_corrected"],
    ) == (2020, 2, "")
    assert calendar_invalid["birth_date_status"] != ""

    assert (
        future_birth["birth_year_corrected"],
        future_birth["birth_month_corrected"],
        future_birth["birth_day_corrected"],
    ) == (2026, 5, 12)
    assert future_birth["birth_date_status"] != ""

    assert (
        before_1906["birth_year_corrected"],
        before_1906["birth_month_corrected"],
        before_1906["birth_day_corrected"],
    ) == (1900, 1, 1)
    assert before_1906["birth_date_status"] != ""

    assert (
        late_entry["entry_year_corrected"],
        late_entry["entry_month_corrected"],
        late_entry["entry_day_corrected"],
    ) == (2026, 1, 1)
    assert late_entry["entry_date_status"] != ""
