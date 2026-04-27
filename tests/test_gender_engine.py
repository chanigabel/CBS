"""Tests for GenderEngine.normalize_gender — valid mappings and invalid fallback."""

import pytest
from src.excel_normalization.engines.gender_engine import GenderEngine


@pytest.fixture
def engine():
    return GenderEngine()


# ---------------------------------------------------------------------------
# Valid female values → 2
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Valid male values → 1
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Empty / None → preserved (pipeline short-circuits; engine returns 1 for
# direct callers — the pipeline layer handles None/whitespace before calling)
# ---------------------------------------------------------------------------

class TestEmptyAndNone:
    def test_none_returns_1(self, engine):
        """Direct engine call with None returns 1 (pipeline never reaches here for None)."""
        assert engine.normalize_gender(None) == 1

    def test_empty_string_returns_1(self, engine):
        """Direct engine call with empty string returns 1."""
        assert engine.normalize_gender("") == 1

    def test_whitespace_only_returns_1(self, engine):
        """Direct engine call with whitespace-only returns 1 (stripped → empty)."""
        assert engine.normalize_gender("   ") == 1


# ---------------------------------------------------------------------------
# Invalid / unrecognized values → "" (empty string)
# ---------------------------------------------------------------------------

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
        """Invalid value must never appear as-is in the corrected field."""
        result = engine.normalize_gender("8")
        assert result != "8", "Raw invalid value must not be copied to corrected field"
        assert result != 8

    def test_invalid_text_does_not_copy_raw_value(self, engine):
        result = engine.normalize_gender("unknown")
        assert result != "unknown"


# ---------------------------------------------------------------------------
# Pipeline integration: apply_gender_normalization with invalid values
# ---------------------------------------------------------------------------

class TestPipelineInvalidGender:
    def _make_pipeline(self):
        from src.excel_normalization.processing.normalization_pipeline import NormalizationPipeline
        return NormalizationPipeline(gender_engine=GenderEngine())

    def test_invalid_numeric_8_corrected_empty(self):
        pipeline = self._make_pipeline()
        row = {"gender": "8"}
        pipeline.apply_gender_normalization(row)
        assert row["gender_corrected"] == "", (
            f"Expected empty corrected for '8', got {row['gender_corrected']!r}"
        )

    def test_invalid_text_corrected_empty(self):
        pipeline = self._make_pipeline()
        row = {"gender": "xyz"}
        pipeline.apply_gender_normalization(row)
        assert row["gender_corrected"] == ""

    def test_valid_1_still_maps_to_1(self):
        pipeline = self._make_pipeline()
        row = {"gender": "1"}
        pipeline.apply_gender_normalization(row)
        assert row["gender_corrected"] == 1

    def test_valid_2_still_maps_to_2(self):
        pipeline = self._make_pipeline()
        row = {"gender": "2"}
        pipeline.apply_gender_normalization(row)
        assert row["gender_corrected"] == 2

    def test_none_preserved_by_pipeline(self):
        pipeline = self._make_pipeline()
        row = {"gender": None}
        pipeline.apply_gender_normalization(row)
        assert row["gender_corrected"] is None

    def test_empty_string_preserved_by_pipeline(self):
        pipeline = self._make_pipeline()
        row = {"gender": ""}
        pipeline.apply_gender_normalization(row)
        assert row["gender_corrected"] == ""

    def test_whitespace_only_preserved_by_pipeline(self):
        """Whitespace-only is caught by the pipeline short-circuit, not the engine."""
        pipeline = self._make_pipeline()
        row = {"gender": "   "}
        pipeline.apply_gender_normalization(row)
        # Pipeline preserves the original whitespace-only value (F-04 behavior)
        assert row["gender_corrected"] == "   "

    def test_hebrew_female_still_maps_to_2(self):
        pipeline = self._make_pipeline()
        row = {"gender": "נ"}
        pipeline.apply_gender_normalization(row)
        assert row["gender_corrected"] == 2

    def test_hebrew_male_still_maps_to_1(self):
        pipeline = self._make_pipeline()
        row = {"gender": "ז"}
        pipeline.apply_gender_normalization(row)
        assert row["gender_corrected"] == 1
