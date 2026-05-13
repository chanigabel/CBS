"""Tests for GenderEngine.normalize_gender - valid mappings and empty fallback."""

import pytest
from src.excel_standardization.engines.gender_engine import GenderEngine


@pytest.fixture
def engine():
    return GenderEngine()


class TestFemaleValues:
    def test_numeric_2(self, engine):
        assert engine.normalize_gender("2") == 2

    def test_numeric_2_int(self, engine):
        assert engine.normalize_gender(2) == 2

    def test_hebrew_nun(self, engine):
        assert engine.normalize_gender("נ") == 2

    def test_hebrew_isha(self, engine):
        assert engine.normalize_gender("אישה") == 2

    def test_hebrew_bat(self, engine):
        assert engine.normalize_gender("בת") == 2

    def test_hebrew_nekeva(self, engine):
        assert engine.normalize_gender("נקבה") == 2

    def test_english_female(self, engine):
        assert engine.normalize_gender("female") == 2

    def test_english_female_upper(self, engine):
        assert engine.normalize_gender("FEMALE") == 2

    def test_english_f(self, engine):
        assert engine.normalize_gender("f") == 2

    def test_english_girl(self, engine):
        assert engine.normalize_gender("girl") == 2

    def test_english_woman(self, engine):
        assert engine.normalize_gender("woman") == 2


class TestMaleValues:
    def test_numeric_1(self, engine):
        assert engine.normalize_gender("1") == 1

    def test_numeric_1_int(self, engine):
        assert engine.normalize_gender(1) == 1

    def test_hebrew_zayin(self, engine):
        assert engine.normalize_gender("ז") == 1

    def test_hebrew_zachar(self, engine):
        assert engine.normalize_gender("זכר") == 1

    def test_hebrew_ben(self, engine):
        assert engine.normalize_gender("בן") == 1

    def test_english_male(self, engine):
        assert engine.normalize_gender("male") == 1

    def test_english_male_upper(self, engine):
        assert engine.normalize_gender("MALE") == 1

    def test_english_m(self, engine):
        assert engine.normalize_gender("m") == 1

    def test_english_man(self, engine):
        assert engine.normalize_gender("man") == 1

    def test_english_boy(self, engine):
        assert engine.normalize_gender("boy") == 1


class TestEmptyAndNone:
    def test_none_returns_empty(self, engine):
        assert engine.normalize_gender(None) == ""

    def test_empty_string_returns_empty(self, engine):
        assert engine.normalize_gender("") == ""

    def test_whitespace_only_returns_empty(self, engine):
        assert engine.normalize_gender("   ") == ""


class TestInvalidValues:
    def test_numeric_8(self, engine):
        assert engine.normalize_gender("8") == ""

    def test_numeric_8_int(self, engine):
        assert engine.normalize_gender(8) == ""

    def test_numeric_0(self, engine):
        assert engine.normalize_gender("0") == ""

    def test_numeric_3(self, engine):
        assert engine.normalize_gender("3") == ""

    def test_numeric_99(self, engine):
        assert engine.normalize_gender("99") == ""

    def test_random_text(self, engine):
        assert engine.normalize_gender("xyz") == ""

    def test_hebrew_unrecognized(self, engine):
        assert engine.normalize_gender("לא ידוע") == ""

    def test_question_mark(self, engine):
        assert engine.normalize_gender("?") == ""

    def test_dash(self, engine):
        assert engine.normalize_gender("-") == ""

    def test_na_string(self, engine):
        assert engine.normalize_gender("N/A") == ""

    def test_invalid_does_not_copy_raw_value(self, engine):
        result = engine.normalize_gender("8")
        assert result != "8"
        assert result != 8

    def test_invalid_text_does_not_copy_raw_value(self, engine):
        result = engine.normalize_gender("unknown")
        assert result != "unknown"


class TestPipelineInvalidGender:
    def _make_pipeline(self):
        from src.excel_standardization.processing.standardization_pipeline import StandardizationPipeline
        return StandardizationPipeline(gender_engine=GenderEngine())

    def test_invalid_numeric_8_corrected_empty(self):
        pipeline = self._make_pipeline()
        row = {"gender": "8"}
        pipeline.apply_gender_standardization(row)
        assert row["gender_corrected"] == ""

    def test_invalid_text_corrected_empty(self):
        pipeline = self._make_pipeline()
        row = {"gender": "xyz"}
        pipeline.apply_gender_standardization(row)
        assert row["gender_corrected"] == ""

    def test_valid_1_still_maps_to_1(self):
        pipeline = self._make_pipeline()
        row = {"gender": "1"}
        pipeline.apply_gender_standardization(row)
        assert row["gender_corrected"] == 1

    def test_valid_2_still_maps_to_2(self):
        pipeline = self._make_pipeline()
        row = {"gender": "2"}
        pipeline.apply_gender_standardization(row)
        assert row["gender_corrected"] == 2

    def test_none_preserved_by_pipeline(self):
        pipeline = self._make_pipeline()
        row = {"gender": None}
        pipeline.apply_gender_standardization(row)
        assert row["gender_corrected"] is None

    def test_empty_string_preserved_by_pipeline(self):
        pipeline = self._make_pipeline()
        row = {"gender": ""}
        pipeline.apply_gender_standardization(row)
        assert row["gender_corrected"] == ""

    def test_whitespace_only_normalized_to_empty_by_pipeline(self):
        pipeline = self._make_pipeline()
        row = {"gender": "   "}
        pipeline.apply_gender_standardization(row)
        assert row["gender"] == "   "
        assert row["gender_corrected"] == ""
        assert "gender_status" not in row

    def test_invalid_non_empty_value_writes_hebrew_status(self):
        pipeline = self._make_pipeline()
        row = {"gender": "8"}
        pipeline.apply_gender_standardization(row)
        assert row["gender_corrected"] == ""
        assert row["gender_status"] == "קוד מין לא תקין - חייב להיות 1 (זכר) או 2 (נקבה)"

    def test_hebrew_female_still_maps_to_2(self):
        pipeline = self._make_pipeline()
        row = {"gender": "נ"}
        pipeline.apply_gender_standardization(row)
        assert row["gender_corrected"] == 2

    def test_hebrew_male_still_maps_to_1(self):
        pipeline = self._make_pipeline()
        row = {"gender": "ז"}
        pipeline.apply_gender_standardization(row)
        assert row["gender_corrected"] == 1
