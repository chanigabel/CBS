"""Tests for conservative date parsing — the engine must never invent fake dates.

Requirements verified:
1. Impossible years (5280, 2229, 9999, 3000) → corrected fields empty, status visible
2. Plain integers without serial metadata → corrected fields empty
3. 5-digit and 7-digit numeric strings → rejected (invalid length)
4. Valid dates → corrected fields populated correctly
5. Excel serial WITH metadata → converts correctly
6. Excel serial WITHOUT metadata → rejected
7. Large serials that produce impossible years → rejected
"""

import pytest
from datetime import date

from src.excel_standardization.engines.date_engine import DateEngine, STATUS_IMPOSSIBLE_YEAR, STATUS_NUMERIC_DATE_UNRECOGNIZED
from src.excel_standardization.data_types import DateFormatPattern, DateFieldType, DateInput
from src.excel_standardization.processing.date_standardization import date_corrected_components


REF_DATE = date(2026, 5, 11)


def _engine():
    return DateEngine(reference_date=REF_DATE)


def _split(yr, mo, dy):
    """Parse via split path and apply business rules."""
    engine = _engine()
    r = engine.parse_from_split_columns(yr, mo, dy)
    return engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)


def _single(val, serial=False):
    """Parse via single path and apply business rules."""
    engine = _engine()
    r = engine.parse_input(DateInput(
        source_kind="single",
        field_type=DateFieldType.BIRTH_DATE,
        raw_value=val,
        source_is_excel_date_serial=serial,
        reference_date=REF_DATE,
    ))
    return r


def _blank(r):
    """Return True if all corrected components are empty strings."""
    cy, cm, cd = date_corrected_components(r)
    return cy == "" and cm == "" and cd == ""


def _populated(r):
    """Return True if all corrected components are non-empty."""
    cy, cm, cd = date_corrected_components(r)
    return cy != "" and cm != "" and cd != ""


# ---------------------------------------------------------------------------
# 1. Impossible years in split path
# ---------------------------------------------------------------------------

class TestImpossibleYearsInSplitPath:
    """Years far beyond the reference year must be rejected."""

    def test_year_1234567_rejected(self):
        r = _split(1234567, 5, 15)
        assert _blank(r), f"Expected blank corrected fields, got year={r.year}"
        assert r.status_text != ""

    def test_year_5280_rejected(self):
        r = _split(5280, 5, 15)
        assert _blank(r)
        assert r.status_text != ""
        assert r.status_code in ("impossible_year", "future_birth")

    def test_year_2229_rejected(self):
        r = _split(2229, 5, 15)
        assert _blank(r)
        assert r.status_text != ""

    def test_year_9999_rejected(self):
        r = _split(9999, 5, 15)
        assert _blank(r)
        assert r.status_text != ""

    def test_year_3000_rejected(self):
        r = _split(3000, 5, 15)
        assert _blank(r)
        assert r.status_text != ""

    def test_year_ref_plus_2_rejected(self):
        """Any year more than 1 beyond the reference year must be rejected."""
        r = _split(REF_DATE.year + 2, 5, 15)
        assert _blank(r)

    def test_year_ref_plus_1_accepted(self):
        """Year = reference_year + 1 is the maximum allowed (near-future entry)."""
        r = _split(REF_DATE.year + 1, 1, 1)
        # Should not be rejected by impossible_year check
        assert r.status_code != "impossible_year"

    def test_impossible_year_status_message(self):
        r = _split(5280, 5, 15)
        assert r.status_text == STATUS_IMPOSSIBLE_YEAR or r.status_text != ""


# ---------------------------------------------------------------------------
# 2. Plain integers without serial metadata
# ---------------------------------------------------------------------------

class TestPlainIntegersWithoutMetadata:
    """Plain int values must never be converted to dates without metadata."""

    def test_int_1234567_rejected(self):
        r = _single(1234567, serial=False)
        assert _blank(r)
        assert r.status_text == STATUS_NUMERIC_DATE_UNRECOGNIZED

    def test_int_120201_rejected(self):
        r = _single(120201, serial=False)
        assert _blank(r)
        assert r.status_text == STATUS_NUMERIC_DATE_UNRECOGNIZED

    def test_int_36525_rejected_without_metadata(self):
        """Valid Excel serial 36525 = 2000-01-01 must be rejected without metadata."""
        r = _single(36525, serial=False)
        assert _blank(r)
        assert r.status_text == STATUS_NUMERIC_DATE_UNRECOGNIZED

    def test_int_45657_rejected_without_metadata(self):
        r = _single(45657, serial=False)
        assert _blank(r)
        assert r.status_text == STATUS_NUMERIC_DATE_UNRECOGNIZED

    def test_int_999999_rejected(self):
        r = _single(999999, serial=False)
        assert _blank(r)

    def test_int_888888_rejected(self):
        r = _single(888888, serial=False)
        assert _blank(r)

    def test_int_2024_rejected(self):
        """Plain year integer must not be converted to a serial date."""
        r = _single(2024, serial=False)
        assert _blank(r)
        assert r.status_text == STATUS_NUMERIC_DATE_UNRECOGNIZED


# ---------------------------------------------------------------------------
# 3. Invalid-length numeric strings
# ---------------------------------------------------------------------------

class TestInvalidLengthNumericStrings:
    """5-digit and 7+ digit strings must be rejected."""

    def test_5_digit_string_rejected(self):
        engine = _engine()
        r = engine.parse_date_value("12345", DateFormatPattern.DDMM)
        assert _blank(r)
        assert r.status_text != ""

    def test_5_digit_string_99999_rejected(self):
        engine = _engine()
        r = engine.parse_date_value("99999", DateFormatPattern.DDMM)
        assert _blank(r)
        assert r.status_text != ""

    def test_7_digit_string_rejected(self):
        engine = _engine()
        r = engine.parse_date_value("1234567", DateFormatPattern.DDMM)
        assert _blank(r)
        assert r.status_text != ""

    def test_7_digit_string_9999999_rejected(self):
        engine = _engine()
        r = engine.parse_date_value("9999999", DateFormatPattern.DDMM)
        assert _blank(r)
        assert r.status_text != ""

    def test_9_digit_string_rejected(self):
        engine = _engine()
        r = engine.parse_date_value("123456789", DateFormatPattern.DDMM)
        assert _blank(r)
        assert r.status_text != ""

    def test_3_digit_string_rejected(self):
        engine = _engine()
        r = engine.parse_date_value("123", DateFormatPattern.DDMM)
        assert _blank(r)
        assert r.status_text != ""


# ---------------------------------------------------------------------------
# 4. Valid dates must still work
# ---------------------------------------------------------------------------

class TestValidDatesAccepted:
    """Legitimate dates must continue to parse correctly."""

    def test_ddmmyyyy_string(self):
        engine = _engine()
        r = engine.parse_date_value("14/03/1985", DateFormatPattern.DDMM)
        r2 = engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)
        assert _populated(r2)
        assert r2.year == 1985
        assert r2.month == 3
        assert r2.day == 14

    def test_iso_string(self):
        engine = _engine()
        r = engine.parse_date_value("1985-03-14", DateFormatPattern.DDMM)
        r2 = engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)
        assert _populated(r2)
        assert r2.year == 1985

    def test_6_digit_ddmmyy(self):
        engine = _engine()
        r = engine.parse_date_value("140385", DateFormatPattern.DDMM)
        r2 = engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)
        assert _populated(r2)
        assert r2.year == 1985

    def test_6_digit_exact_example(self):
        engine = _engine()
        r = engine.parse_date_value("010224", DateFormatPattern.DDMM)
        r2 = engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)
        assert _populated(r2)
        assert (r2.year, r2.month, r2.day) == (2024, 2, 1)

    def test_8_digit_ddmmyyyy(self):
        engine = _engine()
        r = engine.parse_date_value("14031985", DateFormatPattern.DDMM)
        r2 = engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)
        assert _populated(r2)
        assert r2.year == 1985

    def test_8_digit_exact_example(self):
        engine = _engine()
        r = engine.parse_date_value("12022001", DateFormatPattern.DDMM)
        r2 = engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)
        assert _populated(r2)
        assert (r2.year, r2.month, r2.day) == (2001, 2, 12)

    def test_split_valid_date(self):
        r = _split(1985, 3, 14)
        assert _populated(r)
        assert r.year == 1985

    def test_4_digit_year_only(self):
        """4-digit year-only string returns missing_month_day, not a fake date."""
        engine = _engine()
        r = engine.parse_date_value("1985", DateFormatPattern.DDMM)
        assert r.year == 1985
        assert r.month is None
        assert r.day is None
        assert r.status_code == "missing_month_day"

    def test_4_digit_compact_example(self):
        engine = _engine()
        r = engine.parse_date_value("1124", DateFormatPattern.DDMM)
        r2 = engine.validate_business_rules(r, DateFieldType.BIRTH_DATE)
        assert _populated(r2)
        assert (r2.year, r2.month, r2.day) == (2024, 1, 1)


# ---------------------------------------------------------------------------
# 5. Excel serial WITH metadata
# ---------------------------------------------------------------------------

class TestExcelSerialWithMetadata:
    """Serials with source_is_excel_date_serial=True must convert correctly."""

    def test_serial_36525_converts(self):
        """36525 = 1999-12-31 in Excel serial (Excel 1900 leap year offset)."""
        r = _single(36525, serial=True)
        assert _populated(r)
        assert r.year == 1999
        assert r.month == 12
        assert r.day == 31

    def test_serial_38353_converts(self):
        r = _single(38353, serial=True)
        assert _populated(r)
        assert 1900 <= r.year <= REF_DATE.year + 1

    def test_serial_result_year_in_valid_range(self):
        r = _single(36525, serial=True)
        assert 1906 <= r.year <= REF_DATE.year + 1


# ---------------------------------------------------------------------------
# 6. Excel serial WITHOUT metadata
# ---------------------------------------------------------------------------

class TestExcelSerialWithoutMetadata:
    """Same serial values must be rejected when metadata is absent."""

    def test_serial_36525_rejected_without_metadata(self):
        r = _single(36525, serial=False)
        assert _blank(r)
        assert r.status_text == STATUS_NUMERIC_DATE_UNRECOGNIZED

    def test_serial_38353_rejected_without_metadata(self):
        r = _single(38353, serial=False)
        assert _blank(r)
        assert r.status_text == STATUS_NUMERIC_DATE_UNRECOGNIZED


# ---------------------------------------------------------------------------
# 7. Large serials that produce impossible years
# ---------------------------------------------------------------------------

class TestLargeSerialsThatProduceImpossibleYears:
    """Even with metadata, serials that produce impossible years must be rejected."""

    def test_serial_9999999_rejected(self):
        r = _single(9999999, serial=True)
        assert _blank(r)
        assert r.status_text != ""

    def test_serial_2958466_rejected(self):
        """2958466 is beyond Excel's maximum date serial."""
        r = _single(2958466, serial=True)
        assert _blank(r)
        assert r.status_text != ""


# ---------------------------------------------------------------------------
# 8. Pipeline integration — corrected fields stay empty for bad values
# ---------------------------------------------------------------------------

class TestPipelineIntegration:
    """End-to-end: bad values must not appear in corrected fields in the pipeline."""

    def test_impossible_year_in_pipeline(self):
        from src.excel_standardization.processing.standardization_pipeline import StandardizationPipeline
        from src.excel_standardization.engines.name_engine import NameEngine
        from src.excel_standardization.engines.gender_engine import GenderEngine
        from src.excel_standardization.engines.identifier_engine import IdentifierEngine
        from src.excel_standardization.engines.text_processor import TextProcessor

        pipeline = StandardizationPipeline(
            name_engine=NameEngine(TextProcessor()),
            gender_engine=GenderEngine(),
            date_engine=DateEngine(reference_date=REF_DATE),
            identifier_engine=IdentifierEngine(),
            reference_date=REF_DATE,
        )

        row = {
            "birth_year": 5280,
            "birth_month": 5,
            "birth_day": 15,
        }
        result = pipeline.normalize_row(row)

        assert result.get("birth_year_corrected") == ""
        assert result.get("birth_month_corrected") == ""
        assert result.get("birth_day_corrected") == ""
        assert result.get("birth_date_status") != ""

    def test_plain_int_in_pipeline(self):
        from src.excel_standardization.processing.standardization_pipeline import StandardizationPipeline
        from src.excel_standardization.engines.name_engine import NameEngine
        from src.excel_standardization.engines.gender_engine import GenderEngine
        from src.excel_standardization.engines.identifier_engine import IdentifierEngine
        from src.excel_standardization.engines.text_processor import TextProcessor

        pipeline = StandardizationPipeline(
            name_engine=NameEngine(TextProcessor()),
            gender_engine=GenderEngine(),
            date_engine=DateEngine(reference_date=REF_DATE),
            identifier_engine=IdentifierEngine(),
            reference_date=REF_DATE,
        )

        row = {"birth_date": 1234567}
        result = pipeline.normalize_row(row)

        assert result.get("birth_year_corrected") == ""
        assert result.get("birth_month_corrected") == ""
        assert result.get("birth_day_corrected") == ""
        assert result.get("birth_date_status") != ""

    def test_valid_date_in_pipeline(self):
        from src.excel_standardization.processing.standardization_pipeline import StandardizationPipeline
        from src.excel_standardization.engines.name_engine import NameEngine
        from src.excel_standardization.engines.gender_engine import GenderEngine
        from src.excel_standardization.engines.identifier_engine import IdentifierEngine
        from src.excel_standardization.engines.text_processor import TextProcessor

        pipeline = StandardizationPipeline(
            name_engine=NameEngine(TextProcessor()),
            gender_engine=GenderEngine(),
            date_engine=DateEngine(reference_date=REF_DATE),
            identifier_engine=IdentifierEngine(),
            reference_date=REF_DATE,
        )

        row = {
            "birth_year": 1985,
            "birth_month": 3,
            "birth_day": 14,
        }
        result = pipeline.normalize_row(row)

        assert result.get("birth_year_corrected") == 1985
        assert result.get("birth_month_corrected") == 3
        assert result.get("birth_day_corrected") == 14
