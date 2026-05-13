"""Tests for InstitutionReportValidator.

Covers all mandatory field validations for institution-report files:
    - MosadID, SugMosad, MisparDiraBeMosad
    - ShemPrati, ShemMishpaha
    - MisparZehut (required, duplicate within sheet, duplicate across workbook)
    - Min (gender)
    - Birth date (ShnatLida, HodeshLida, YomLida)
    - Entry date (shnatknisa, Hodeshknisa, YomKnisa)
    - Cross-sheet duplicate detection
    - Sheet-specific rules (YomKnisa required only for DayarimYahidim)
"""

import pytest
from src.excel_standardization.validation.institution_report_validator import (
    InstitutionReportValidator,
    RowValidationResult,
    SHEET_ANASHEY_TZEVET,
    SHEET_DAYARIM_YAHIDIM,
    SHEET_MESHKEY_BAYT,
    MSG_MOSAD_ID_MISSING,
    MSG_SUG_MOSAD_MISSING,
    MSG_SUG_MOSAD_NOT_NUMERIC,
    MSG_SUG_MOSAD_TOO_SHORT,
    MSG_DIRA_NOT_NUMERIC,
    MSG_SHEM_PRATI_MISSING,
    MSG_SHEM_MISHPAHA_MISSING,
    MSG_MISPAR_ZEHUT_MISSING,
    MSG_MISPAR_ZEHUT_DUPLICATE_SHEET,
    MSG_MISPAR_ZEHUT_DUPLICATE_WORKBOOK,
    MSG_MIN_INVALID,
    MSG_SHNAT_LIDA_MISSING,
    MSG_SHNAT_LIDA_NOT_NUMERIC,
    MSG_SHNAT_LIDA_TOO_EARLY,
    MSG_SHNAT_LIDA_FUTURE,
    MSG_HODESH_LIDA_MISSING,
    MSG_HODESH_LIDA_NOT_NUMERIC,
    MSG_HODESH_LIDA_RANGE,
    MSG_YOM_LIDA_MISSING,
    MSG_YOM_LIDA_NOT_NUMERIC,
    MSG_YOM_LIDA_RANGE,
    MSG_SHNAT_KNISA_MISSING,
    MSG_SHNAT_KNISA_NOT_NUMERIC,
    MSG_SHNAT_KNISA_AFTER_CENSUS,
    MSG_HODESH_KNISA_MISSING,
    MSG_HODESH_KNISA_NOT_NUMERIC,
    MSG_HODESH_KNISA_RANGE,
    MSG_YOM_KNISA_MISSING,
    MSG_YOM_KNISA_NOT_NUMERIC,
    MSG_YOM_KNISA_RANGE,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_valid_row(**overrides) -> dict:
    """Return a minimal valid row for DayarimYahidim."""
    row = {
        "MosadID": "12345",
        "SugMosad": "100",
        "first_name_corrected": "יוסי",
        "last_name_corrected": "כהן",
        "id_number_corrected": "039337423",
        "gender_corrected": 1,
        "birth_year_corrected": 1980,
        "birth_month_corrected": 5,
        "birth_day_corrected": 15,
        "entry_year_corrected": 2010,
        "entry_month_corrected": 3,
        "entry_day_corrected": 1,
    }
    row.update(overrides)
    return row


def _validator(sheet=SHEET_DAYARIM_YAHIDIM, census_year=2025) -> InstitutionReportValidator:
    return InstitutionReportValidator(sheet_name=sheet, census_year=census_year)


def _messages(result: RowValidationResult) -> list:
    return [f.message for f in result.findings]


# ---------------------------------------------------------------------------
# MosadID
# ---------------------------------------------------------------------------

class TestMosadID:
    def test_valid(self):
        row = _make_valid_row()
        r = _validator().validate_row(row)
        assert MSG_MOSAD_ID_MISSING not in _messages(r)

    def test_missing(self):
        row = _make_valid_row(MosadID=None)
        r = _validator().validate_row(row)
        assert MSG_MOSAD_ID_MISSING in _messages(r)

    def test_empty_string(self):
        row = _make_valid_row(MosadID="")
        r = _validator().validate_row(row)
        assert MSG_MOSAD_ID_MISSING in _messages(r)

    def test_non_numeric_allowed(self):
        row = _make_valid_row(MosadID="abc")
        r = _validator().validate_row(row)
        assert MSG_MOSAD_ID_MISSING not in _messages(r)

    def test_short_value_allowed(self):
        row = _make_valid_row(MosadID="12")
        r = _validator().validate_row(row)
        assert MSG_MOSAD_ID_MISSING not in _messages(r)

    def test_exactly_3_digits_ok(self):
        row = _make_valid_row(MosadID="123")
        r = _validator().validate_row(row)
        assert MSG_MOSAD_ID_MISSING not in _messages(r)


# ---------------------------------------------------------------------------
# SugMosad
# ---------------------------------------------------------------------------

class TestSugMosad:
    def test_valid(self):
        row = _make_valid_row()
        r = _validator().validate_row(row)
        assert MSG_SUG_MOSAD_MISSING not in _messages(r)

    def test_missing(self):
        row = _make_valid_row(SugMosad=None)
        r = _validator().validate_row(row)
        assert MSG_SUG_MOSAD_MISSING in _messages(r)

    def test_not_numeric(self):
        row = _make_valid_row(SugMosad="abc")
        r = _validator().validate_row(row)
        assert MSG_SUG_MOSAD_NOT_NUMERIC in _messages(r)

    def test_too_short(self):
        row = _make_valid_row(SugMosad="10")
        r = _validator().validate_row(row)
        assert MSG_SUG_MOSAD_TOO_SHORT in _messages(r)

    def test_exactly_3_digits_ok(self):
        row = _make_valid_row(SugMosad="100")
        r = _validator().validate_row(row)
        assert MSG_SUG_MOSAD_TOO_SHORT not in _messages(r)


# ---------------------------------------------------------------------------
# MisparDiraBeMosad
# ---------------------------------------------------------------------------

class TestMisparDiraBeMosad:
    def test_optional_empty_ok(self):
        row = _make_valid_row()
        # No MisparDiraBeMosad key at all
        r = InstitutionReportValidator(
            sheet_name=SHEET_MESHKEY_BAYT, census_year=2025
        ).validate_row(row)
        assert MSG_DIRA_NOT_NUMERIC not in _messages(r)

    def test_numeric_ok(self):
        row = _make_valid_row(MisparDiraBeMosad="42")
        r = InstitutionReportValidator(
            sheet_name=SHEET_MESHKEY_BAYT, census_year=2025
        ).validate_row(row)
        assert MSG_DIRA_NOT_NUMERIC not in _messages(r)

    def test_not_numeric(self):
        row = _make_valid_row(MisparDiraBeMosad="abc")
        r = InstitutionReportValidator(
            sheet_name=SHEET_MESHKEY_BAYT, census_year=2025
        ).validate_row(row)
        assert MSG_DIRA_NOT_NUMERIC in _messages(r)

    def test_skipped_for_dayarim(self):
        """DayarimYahidim does not have MisparDiraBeMosad — should not flag it."""
        row = _make_valid_row(MisparDiraBeMosad="abc")
        r = _validator(sheet=SHEET_DAYARIM_YAHIDIM).validate_row(row)
        assert MSG_DIRA_NOT_NUMERIC not in _messages(r)


# ---------------------------------------------------------------------------
# ShemPrati / ShemMishpaha
# ---------------------------------------------------------------------------

class TestNames:
    def test_shem_prati_missing(self):
        row = _make_valid_row()
        row.pop("first_name_corrected", None)
        r = _validator().validate_row(row)
        assert MSG_SHEM_PRATI_MISSING in _messages(r)

    def test_shem_prati_empty(self):
        row = _make_valid_row(**{"first_name_corrected": ""})
        r = _validator().validate_row(row)
        assert MSG_SHEM_PRATI_MISSING in _messages(r)

    def test_shem_prati_ok(self):
        row = _make_valid_row()
        r = _validator().validate_row(row)
        assert MSG_SHEM_PRATI_MISSING not in _messages(r)

    def test_shem_mishpaha_missing(self):
        row = _make_valid_row()
        row.pop("last_name_corrected", None)
        r = _validator().validate_row(row)
        assert MSG_SHEM_MISHPAHA_MISSING in _messages(r)

    def test_shem_mishpaha_empty(self):
        row = _make_valid_row(**{"last_name_corrected": ""})
        r = _validator().validate_row(row)
        assert MSG_SHEM_MISHPAHA_MISSING in _messages(r)

    def test_falls_back_to_original(self):
        """Validator should fall back to original field when corrected is absent."""
        row = _make_valid_row()
        row.pop("first_name_corrected", None)
        row["first_name"] = "שרה"
        r = _validator().validate_row(row)
        assert MSG_SHEM_PRATI_MISSING not in _messages(r)

    def test_mixed_corrected_and_original_selection(self):
        row = _make_valid_row(
            first_name="Original First",
            first_name_corrected="Corrected First",
            last_name="Original Last",
            last_name_corrected="",
        )
        r = _validator().validate_row(row)
        assert MSG_SHEM_PRATI_MISSING not in _messages(r)
        assert MSG_SHEM_MISHPAHA_MISSING not in _messages(r)

    def test_unexpected_corrected_field_type_is_stringified_for_validation(self):
        row = _make_valid_row(
            first_name_corrected=["Corrected"],
            last_name_corrected={"name": "Last"},
        )
        r = _validator().validate_row(row)
        assert MSG_SHEM_PRATI_MISSING not in _messages(r)
        assert MSG_SHEM_MISHPAHA_MISSING not in _messages(r)


# ---------------------------------------------------------------------------
# MisparZehut
# ---------------------------------------------------------------------------

class TestMisparZehut:
    def test_missing(self):
        row = _make_valid_row()
        row.pop("id_number_corrected", None)
        r = _validator().validate_row(row)
        assert MSG_MISPAR_ZEHUT_MISSING in _messages(r)

    def test_empty(self):
        row = _make_valid_row(**{"id_number_corrected": ""})
        r = _validator().validate_row(row)
        assert MSG_MISPAR_ZEHUT_MISSING in _messages(r)

    def test_present_ok(self):
        row = _make_valid_row()
        r = _validator().validate_row(row)
        assert MSG_MISPAR_ZEHUT_MISSING not in _messages(r)

    def test_duplicate_within_sheet(self):
        rows = [
            _make_valid_row(**{"id_number_corrected": "039337423", "_row_uid": "r1"}),
            _make_valid_row(**{"id_number_corrected": "039337423", "_row_uid": "r2"}),
        ]
        v = _validator()
        results = v.validate_sheet(rows)
        # First occurrence: no duplicate error
        assert MSG_MISPAR_ZEHUT_DUPLICATE_SHEET not in _messages(results[0])
        # Second occurrence: duplicate error
        assert MSG_MISPAR_ZEHUT_DUPLICATE_SHEET in _messages(results[1])

    def test_no_duplicate_different_ids(self):
        rows = [
            _make_valid_row(**{"id_number_corrected": "039337423"}),
            _make_valid_row(**{"id_number_corrected": "000000018"}),
        ]
        v = _validator()
        results = v.validate_sheet(rows)
        assert MSG_MISPAR_ZEHUT_DUPLICATE_SHEET not in _messages(results[0])
        assert MSG_MISPAR_ZEHUT_DUPLICATE_SHEET not in _messages(results[1])

    def test_duplicate_across_workbook(self):
        shared_id = "039337423"
        sheets = {
            SHEET_DAYARIM_YAHIDIM: [_make_valid_row(**{"id_number_corrected": shared_id})],
            SHEET_ANASHEY_TZEVET: [_make_valid_row(**{"id_number_corrected": shared_id})],
        }
        v = InstitutionReportValidator(census_year=2025)
        results = v.validate_workbook(sheets)
        # Both rows should have the cross-workbook warning
        dayarim_msgs = _messages(results[SHEET_DAYARIM_YAHIDIM][0])
        anashey_msgs = _messages(results[SHEET_ANASHEY_TZEVET][0])
        assert MSG_MISPAR_ZEHUT_DUPLICATE_WORKBOOK in dayarim_msgs
        assert MSG_MISPAR_ZEHUT_DUPLICATE_WORKBOOK in anashey_msgs

    def test_no_cross_workbook_duplicate_unique_ids(self):
        sheets = {
            SHEET_DAYARIM_YAHIDIM: [_make_valid_row(**{"id_number_corrected": "039337423"})],
            SHEET_ANASHEY_TZEVET: [_make_valid_row(**{"id_number_corrected": "000000018"})],
        }
        v = InstitutionReportValidator(census_year=2025)
        results = v.validate_workbook(sheets)
        assert MSG_MISPAR_ZEHUT_DUPLICATE_WORKBOOK not in _messages(results[SHEET_DAYARIM_YAHIDIM][0])
        assert MSG_MISPAR_ZEHUT_DUPLICATE_WORKBOOK not in _messages(results[SHEET_ANASHEY_TZEVET][0])

    def test_empty_corrected_id_not_flagged_as_duplicate(self):
        """Empty corrected ID (invalid/rejected by engine) must not be flagged as duplicate."""
        rows = [
            _make_valid_row(**{"id_number_corrected": ""}),
            _make_valid_row(**{"id_number_corrected": ""}),
        ]
        v = _validator()
        results = v.validate_sheet(rows)
        # Both rows have empty corrected ID — should get MISSING error, not DUPLICATE
        assert MSG_MISPAR_ZEHUT_MISSING in _messages(results[0])
        assert MSG_MISPAR_ZEHUT_MISSING in _messages(results[1])
        assert MSG_MISPAR_ZEHUT_DUPLICATE_SHEET not in _messages(results[0])
        assert MSG_MISPAR_ZEHUT_DUPLICATE_SHEET not in _messages(results[1])


# ---------------------------------------------------------------------------
# Min (gender)
# ---------------------------------------------------------------------------

class TestMin:
    def test_valid_1(self):
        row = _make_valid_row(**{"gender_corrected": 1})
        r = _validator().validate_row(row)
        assert MSG_MIN_INVALID not in _messages(r)

    def test_valid_2(self):
        row = _make_valid_row(**{"gender_corrected": 2})
        r = _validator().validate_row(row)
        assert MSG_MIN_INVALID not in _messages(r)

    def test_invalid_code(self):
        row = _make_valid_row(**{"gender_corrected": 3})
        r = _validator().validate_row(row)
        assert MSG_MIN_INVALID in _messages(r)

    def test_empty_string_from_engine(self):
        """GenderEngine returns '' for unrecognized values — should flag as invalid."""
        row = _make_valid_row(**{"gender_corrected": ""})
        r = _validator().validate_row(row)
        assert MSG_MIN_INVALID in _messages(r)

    def test_no_gender_field_ok(self):
        """Gender is optional — missing field should not produce an error."""
        row = _make_valid_row()
        row.pop("gender_corrected", None)
        row.pop("gender", None)
        r = _validator().validate_row(row)
        assert MSG_MIN_INVALID not in _messages(r)


# ---------------------------------------------------------------------------
# Birth date
# ---------------------------------------------------------------------------

class TestBirthDate:
    def test_valid(self):
        row = _make_valid_row()
        r = _validator().validate_row(row)
        for msg in [MSG_SHNAT_LIDA_MISSING, MSG_HODESH_LIDA_MISSING, MSG_YOM_LIDA_MISSING]:
            assert msg not in _messages(r)

    def test_year_missing(self):
        row = _make_valid_row()
        row.pop("birth_year_corrected", None)
        r = _validator().validate_row(row)
        assert MSG_SHNAT_LIDA_MISSING in _messages(r)

    def test_year_not_numeric(self):
        row = _make_valid_row(**{"birth_year_corrected": "abc"})
        r = _validator().validate_row(row)
        assert MSG_SHNAT_LIDA_NOT_NUMERIC in _messages(r)

    def test_year_before_1906(self):
        row = _make_valid_row(**{"birth_year_corrected": 1905})
        r = _validator().validate_row(row)
        assert MSG_SHNAT_LIDA_TOO_EARLY in _messages(r)

    def test_year_1906_ok(self):
        row = _make_valid_row(**{"birth_year_corrected": 1906})
        r = _validator().validate_row(row)
        assert MSG_SHNAT_LIDA_TOO_EARLY not in _messages(r)

    def test_year_future(self):
        from datetime import date
        future_year = date.today().year + 1
        row = _make_valid_row(**{"birth_year_corrected": future_year})
        r = _validator().validate_row(row)
        assert MSG_SHNAT_LIDA_FUTURE in _messages(r)

    def test_month_missing(self):
        row = _make_valid_row()
        row.pop("birth_month_corrected", None)
        r = _validator().validate_row(row)
        assert MSG_HODESH_LIDA_MISSING in _messages(r)

    def test_month_not_numeric(self):
        row = _make_valid_row(**{"birth_month_corrected": "abc"})
        r = _validator().validate_row(row)
        assert MSG_HODESH_LIDA_NOT_NUMERIC in _messages(r)

    def test_month_out_of_range(self):
        row = _make_valid_row(**{"birth_month_corrected": 13})
        r = _validator().validate_row(row)
        assert MSG_HODESH_LIDA_RANGE in _messages(r)

    def test_month_zero(self):
        row = _make_valid_row(**{"birth_month_corrected": 0})
        r = _validator().validate_row(row)
        assert MSG_HODESH_LIDA_RANGE in _messages(r)

    def test_day_missing(self):
        row = _make_valid_row()
        row.pop("birth_day_corrected", None)
        r = _validator().validate_row(row)
        assert MSG_YOM_LIDA_MISSING in _messages(r)

    def test_day_not_numeric(self):
        row = _make_valid_row(**{"birth_day_corrected": "abc"})
        r = _validator().validate_row(row)
        assert MSG_YOM_LIDA_NOT_NUMERIC in _messages(r)

    def test_day_out_of_range(self):
        row = _make_valid_row(**{"birth_day_corrected": 32})
        r = _validator().validate_row(row)
        assert MSG_YOM_LIDA_RANGE in _messages(r)


# ---------------------------------------------------------------------------
# Entry date
# ---------------------------------------------------------------------------

class TestEntryDate:
    def test_valid(self):
        row = _make_valid_row()
        r = _validator().validate_row(row)
        for msg in [MSG_SHNAT_KNISA_MISSING, MSG_HODESH_KNISA_MISSING]:
            assert msg not in _messages(r)

    def test_year_missing(self):
        row = _make_valid_row()
        row.pop("entry_year_corrected", None)
        r = _validator().validate_row(row)
        assert MSG_SHNAT_KNISA_MISSING in _messages(r)

    def test_year_not_numeric(self):
        row = _make_valid_row(**{"entry_year_corrected": "abc"})
        r = _validator().validate_row(row)
        assert MSG_SHNAT_KNISA_NOT_NUMERIC in _messages(r)

    def test_year_after_census(self):
        row = _make_valid_row(**{"entry_year_corrected": 2026})
        r = _validator(census_year=2025).validate_row(row)
        assert MSG_SHNAT_KNISA_AFTER_CENSUS in _messages(r)

    def test_year_equals_census_ok(self):
        row = _make_valid_row(**{"entry_year_corrected": 2025})
        r = _validator(census_year=2025).validate_row(row)
        assert MSG_SHNAT_KNISA_AFTER_CENSUS not in _messages(r)

    def test_month_missing(self):
        row = _make_valid_row()
        row.pop("entry_month_corrected", None)
        r = _validator().validate_row(row)
        assert MSG_HODESH_KNISA_MISSING in _messages(r)

    def test_month_not_numeric(self):
        row = _make_valid_row(**{"entry_month_corrected": "abc"})
        r = _validator().validate_row(row)
        assert MSG_HODESH_KNISA_NOT_NUMERIC in _messages(r)

    def test_month_out_of_range(self):
        row = _make_valid_row(**{"entry_month_corrected": 13})
        r = _validator().validate_row(row)
        assert MSG_HODESH_KNISA_RANGE in _messages(r)

    def test_yom_knisa_required_for_dayarim(self):
        row = _make_valid_row()
        row.pop("entry_day_corrected", None)
        r = _validator(sheet=SHEET_DAYARIM_YAHIDIM).validate_row(row)
        assert MSG_YOM_KNISA_MISSING in _messages(r)

    def test_yom_knisa_optional_for_anashey(self):
        row = _make_valid_row()
        row.pop("entry_day_corrected", None)
        r = _validator(sheet=SHEET_ANASHEY_TZEVET).validate_row(row)
        assert MSG_YOM_KNISA_MISSING not in _messages(r)

    def test_yom_knisa_optional_for_meshkey(self):
        row = _make_valid_row()
        row.pop("entry_day_corrected", None)
        r = _validator(sheet=SHEET_MESHKEY_BAYT).validate_row(row)
        assert MSG_YOM_KNISA_MISSING not in _messages(r)

    def test_yom_knisa_not_numeric(self):
        row = _make_valid_row(**{"entry_day_corrected": "abc"})
        r = _validator(sheet=SHEET_DAYARIM_YAHIDIM).validate_row(row)
        assert MSG_YOM_KNISA_NOT_NUMERIC in _messages(r)

    def test_yom_knisa_out_of_range(self):
        row = _make_valid_row(**{"entry_day_corrected": 32})
        r = _validator(sheet=SHEET_DAYARIM_YAHIDIM).validate_row(row)
        assert MSG_YOM_KNISA_RANGE in _messages(r)


# ---------------------------------------------------------------------------
# Row-level status written back into row dict
# ---------------------------------------------------------------------------

class TestRowStatusWriteback:
    def test_valid_row_has_empty_status(self):
        row = _make_valid_row()
        _validator().validate_row(row)
        assert row["_validation_status"] == ""
        assert row["_validation_ok"] is True

    def test_invalid_row_has_status_message(self):
        row = _make_valid_row(MosadID=None)
        _validator().validate_row(row)
        assert MSG_MOSAD_ID_MISSING in row["_validation_status"]
        assert row["_validation_ok"] is False

    def test_multiple_errors_pipe_separated(self):
        row = _make_valid_row(MosadID=None, SugMosad=None)
        _validator().validate_row(row)
        assert "|" in row["_validation_status"]


# ---------------------------------------------------------------------------
# validate_sheet convenience
# ---------------------------------------------------------------------------

class TestValidateSheet:
    def test_returns_one_result_per_row(self):
        rows = [_make_valid_row() for _ in range(5)]
        v = _validator()
        results = v.validate_sheet(rows)
        assert len(results) == 5

    def test_all_valid_rows(self):
        # Use distinct IDs to avoid duplicate detection
        rows = [
            _make_valid_row(**{"id_number_corrected": f"00000001{i}"}) for i in range(3)
        ]
        v = _validator()
        results = v.validate_sheet(rows)
        assert all(r.is_valid for r in results)


# ---------------------------------------------------------------------------
# Sheet name resolver
# ---------------------------------------------------------------------------

class TestSheetNameResolver:
    def test_hebrew_dayarim(self):
        from src.excel_standardization.services.sheet_name_resolver import resolve_canonical_sheet_name
        assert resolve_canonical_sheet_name("דיירים יחידים") == "DayarimYahidim"

    def test_hebrew_anashey(self):
        from src.excel_standardization.services.sheet_name_resolver import resolve_canonical_sheet_name
        assert resolve_canonical_sheet_name("אנשי צוות ובני משפחותיהם") == "AnasheyTzevet"

    def test_hebrew_meshkey(self):
        from src.excel_standardization.services.sheet_name_resolver import resolve_canonical_sheet_name
        assert resolve_canonical_sheet_name("מתגוררים במשקי בית") == "MeshkeyBayt"

    def test_already_canonical(self):
        from src.excel_standardization.services.sheet_name_resolver import resolve_canonical_sheet_name
        assert resolve_canonical_sheet_name("DayarimYahidim") == "DayarimYahidim"

    def test_unknown_sheet_unchanged(self):
        from src.excel_standardization.services.sheet_name_resolver import resolve_canonical_sheet_name
        assert resolve_canonical_sheet_name("SomeOtherSheet") == "SomeOtherSheet"
