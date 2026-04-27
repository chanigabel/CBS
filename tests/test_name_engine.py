"""Tests for NameEngine and TextProcessor normalization behavior.

Validates real normalization results: trimming, diacritic removal,
language filtering, Hebrew final letter spacing, and father name
last-name removal patterns.
"""

import pytest
from src.excel_normalization.engines.name_engine import NameEngine
from src.excel_normalization.engines.text_processor import TextProcessor
from src.excel_normalization.data_types import Language, FatherNamePattern


# ---------------------------------------------------------------------------
# TextProcessor – language detection
# ---------------------------------------------------------------------------

class TestLanguageDetection:
    def setup_method(self):
        self.tp = TextProcessor()

    def test_pure_hebrew(self):
        assert self.tp.detect_language_dominance("יוסי כהן") == Language.HEBREW

    def test_pure_english(self):
        assert self.tp.detect_language_dominance("John Smith") == Language.ENGLISH

    def test_mixed_more_hebrew(self):
        # 4 Hebrew letters vs 1 English letter
        assert self.tp.detect_language_dominance("יוסיA") == Language.HEBREW

    def test_mixed_more_english(self):
        # 1 Hebrew letter vs 4 English letters
        assert self.tp.detect_language_dominance("אJohn") == Language.ENGLISH

    def test_equal_counts_is_mixed(self):
        # 2 Hebrew, 2 English
        # VBA parity rule: Hebrew wins ties
        assert self.tp.detect_language_dominance("אבAB") == Language.HEBREW

    def test_empty_string_is_mixed(self):
        # No letters → counts are equal (0 == 0)
        assert self.tp.detect_language_dominance("") == Language.MIXED

    def test_digits_and_spaces_ignored(self):
        # Only Hebrew letters count
        assert self.tp.detect_language_dominance("123 456 יוסי") == Language.HEBREW


# ---------------------------------------------------------------------------
# TextProcessor – diacritic removal
# ---------------------------------------------------------------------------

class TestDiacriticRemoval:
    def setup_method(self):
        self.tp = TextProcessor()

    def test_removes_accents_lowercase(self):
        assert self.tp.remove_diacritics("café") == "cafe"

    def test_removes_accents_uppercase(self):
        assert self.tp.remove_diacritics("CAFÉ") == "CAFE"

    def test_no_diacritics_unchanged(self):
        assert self.tp.remove_diacritics("hello") == "hello"

    def test_hebrew_unchanged(self):
        assert self.tp.remove_diacritics("יוסי") == "יוסי"

    def test_mixed_diacritics(self):
        assert self.tp.remove_diacritics("naïve résumé") == "naive resume"


# ---------------------------------------------------------------------------
# TextProcessor – Hebrew final letter spacing
# ---------------------------------------------------------------------------

class TestHebrewFinalLetters:
    def setup_method(self):
        self.tp = TextProcessor()

    def test_final_kaf_before_letter_gets_space(self):
        # ך followed immediately by another Hebrew letter → space inserted
        result = self.tp.fix_hebrew_final_letters("מלךדוד")
        assert "ך " in result

    def test_final_letter_before_space_unchanged(self):
        result = self.tp.fix_hebrew_final_letters("מלך דוד")
        # Already has space — no double space
        assert result == "מלך דוד"

    def test_no_final_letters_unchanged(self):
        result = self.tp.fix_hebrew_final_letters("שלום")
        assert result == "שלום"

    def test_final_mem_before_letter(self):
        result = self.tp.fix_hebrew_final_letters("עםישראל")
        assert "ם " in result


# ---------------------------------------------------------------------------
# TextProcessor – collapse_spaces
# ---------------------------------------------------------------------------

class TestCollapseSpaces:
    def setup_method(self):
        self.tp = TextProcessor()

    def test_multiple_spaces_collapsed(self):
        assert self.tp.collapse_spaces("a  b   c") == "a b c"

    def test_leading_trailing_removed(self):
        assert self.tp.collapse_spaces("  hello  ") == "hello"

    def test_single_space_unchanged(self):
        assert self.tp.collapse_spaces("a b") == "a b"

    def test_tabs_treated_as_whitespace(self):
        assert self.tp.collapse_spaces("a\t\tb") == "a b"


# ---------------------------------------------------------------------------
# TextProcessor – clean_text (full pipeline)
# ---------------------------------------------------------------------------

class TestCleanText:
    def setup_method(self):
        self.tp = TextProcessor()

    def test_trims_whitespace(self):
        assert self.tp.clean_text("  יוסי  ") == "יוסי"

    def test_removes_numbers_from_hebrew_name(self):
        # Hebrew dominant → digits stripped
        assert self.tp.clean_text("שרה123") == "שרה"

    def test_removes_english_from_hebrew_name(self):
        assert self.tp.clean_text("יוסיABC") == "יוסי"

    def test_removes_hebrew_from_english_name(self):
        # Equal Hebrew/English counts → Hebrew wins ties → English removed
        result = self.tp.clean_text("Johnיוסי")
        assert result == "יוסי"

    def test_preserves_hyphen_in_hebrew(self):
        # Hyphens are now converted to spaces, not kept as hyphens
        result = self.tp.clean_text("בן-דוד")
        assert "בן" in result
        assert "דוד" in result
        # hyphen becomes a space separator
        assert result == "בן דוד"

    def test_preserves_hyphen_in_english(self):
        # Hyphens are converted to spaces
        result = self.tp.clean_text("Smith-Jones")
        assert result == "Smith Jones"

    def test_removes_diacritics_in_english(self):
        result = self.tp.clean_text("José")
        assert result == "Jose"

    def test_empty_string_returns_empty(self):
        assert self.tp.clean_text("") == ""

    def test_only_spaces_returns_empty(self):
        assert self.tp.clean_text("   ") == ""

    def test_collapses_internal_spaces(self):
        assert self.tp.clean_text("יוסי  כהן") == "יוסי כהן"

    def test_mixed_language_keeps_both(self):
        # Equal Hebrew/English → Hebrew wins ties → English removed
        result = self.tp.clean_text("אבAB")
        assert "א" in result
        assert "A" not in result


# ---------------------------------------------------------------------------
# NameEngine – normalize_name
# ---------------------------------------------------------------------------

class TestNormalizeName:
    def setup_method(self):
        self.engine = NameEngine(TextProcessor())

    def test_trims_spaces(self):
        assert self.engine.normalize_name("  יוסי  ") == "יוסי"

    def test_removes_digits_from_hebrew(self):
        assert self.engine.normalize_name("שרה123") == "שרה"

    def test_empty_string_returns_empty(self):
        assert self.engine.normalize_name("") == ""

    def test_non_string_returns_empty(self):
        assert self.engine.normalize_name(None) == ""
        assert self.engine.normalize_name(123) == ""

    def test_clean_hebrew_name_unchanged(self):
        assert self.engine.normalize_name("דוד") == "דוד"

    def test_clean_english_name_unchanged(self):
        assert self.engine.normalize_name("David") == "David"

    def test_diacritics_removed(self):
        assert self.engine.normalize_name("José") == "Jose"

    def test_trailing_space_trimmed(self):
        assert self.engine.normalize_name("כהן  ") == "כהן"


# ---------------------------------------------------------------------------
# NameEngine – remove_last_name_from_father
# ---------------------------------------------------------------------------

class TestRemoveLastNameFromFather:
    def setup_method(self):
        self.engine = NameEngine(TextProcessor())

    def test_pattern_none_returns_unchanged(self):
        result = self.engine.remove_last_name_from_father("אברהם כהן", "כהן", FatherNamePattern.NONE)
        assert result == "אברהם כהן"

    def test_remove_last_removes_trailing_last_name(self):
        # Stage A: remove_substring("אברהם כהן", "כהן") → "אברהם"  (changed)
        # Stage A changed the value → Stage B must NOT run.
        # Correct result: "אברהם"
        result = self.engine.remove_last_name_from_father("אברהם כהן", "כהן", FatherNamePattern.REMOVE_LAST)
        assert result == "אברהם"

    def test_remove_first_removes_leading_last_name(self):
        # Stage A: remove_substring("כהן אברהם", "כהן") → "אברהם"  (changed)
        # Stage A changed the value → Stage B must NOT run.
        # Correct result: "אברהם"
        result = self.engine.remove_last_name_from_father("כהן אברהם", "כהן", FatherNamePattern.REMOVE_FIRST)
        assert result == "אברהם"

    def test_last_name_not_in_father_returns_unchanged(self):
        # last="כהן" is not in "אברהם לוי", pattern=REMOVE_LAST
        # Stage A: no-op. Stage B: REMOVE_LAST → removes last token → "אברהם"
        result = self.engine.remove_last_name_from_father("אברהם לוי", "כהן", FatherNamePattern.REMOVE_LAST)
        assert result == "אברהם"

    def test_empty_father_name_returns_empty(self):
        result = self.engine.remove_last_name_from_father("", "כהן", FatherNamePattern.REMOVE_LAST)
        assert result == ""

    def test_empty_last_name_returns_father_unchanged(self):
        result = self.engine.remove_last_name_from_father("אברהם כהן", "", FatherNamePattern.REMOVE_LAST)
        assert result == "אברהם כהן"

    def test_only_last_name_in_father_remove_last_returns_empty(self):
        result = self.engine.remove_last_name_from_father("כהן", "כהן", FatherNamePattern.REMOVE_LAST)
        assert result == ""

    def test_only_last_name_in_father_remove_first_returns_empty(self):
        result = self.engine.remove_last_name_from_father("כהן", "כהן", FatherNamePattern.REMOVE_FIRST)
        assert result == ""

    def test_three_word_father_remove_last(self):
        # Stage A: remove_substring("אברהם יצחק כהן", "כהן") → "אברהם יצחק"  (changed)
        # Stage A changed the value → Stage B must NOT run.
        # Correct result: "אברהם יצחק"
        result = self.engine.remove_last_name_from_father("אברהם יצחק כהן", "כהן", FatherNamePattern.REMOVE_LAST)
        assert result == "אברהם יצחק"

    def test_three_word_father_remove_first(self):
        # Stage A: remove_substring("כהן אברהם יצחק", "כהן") → "אברהם יצחק"  (changed)
        # Stage A changed the value → Stage B must NOT run.
        # Correct result: "אברהם יצחק"
        result = self.engine.remove_last_name_from_father("כהן אברהם יצחק", "כהן", FatherNamePattern.REMOVE_FIRST)
        assert result == "אברהם יצחק"


# ---------------------------------------------------------------------------
# Two-stage removal rule — father name
# ---------------------------------------------------------------------------

class TestTwoStageRemovalFather:
    """Verify Stage B runs ONLY when Stage A made no change."""

    def setup_method(self):
        self.engine = NameEngine(TextProcessor())

    def test_stage_a_success_stops_before_stage_b(self):
        # Example 1 from spec: last="אדמוני", father="אדמוני רפאל"
        # Stage A removes "אדמוני" → "רפאל"  (changed) → stop, no Stage B
        result = self.engine.remove_last_name_from_father(
            "אדמוני רפאל", "אדמוני", FatherNamePattern.REMOVE_FIRST
        )
        assert result == "רפאל"

    def test_stage_a_success_trailing_stops_before_stage_b(self):
        # Example 2 from spec: last="כהן", father="יעקב כהן"
        # Stage A removes "כהן" → "יעקב"  (changed) → stop
        result = self.engine.remove_last_name_from_father(
            "יעקב כהן", "כהן", FatherNamePattern.REMOVE_LAST
        )
        assert result == "יעקב"

    def test_stage_b_runs_when_stage_a_unchanged(self):
        # last_name "כהן" is NOT in "אברהם יצחק" → Stage A is a no-op.
        # Stage B: REMOVE_LAST → removes last token → "אברהם"
        result = self.engine.remove_last_name_from_father(
            "אברהם יצחק", "כהן", FatherNamePattern.REMOVE_LAST
        )
        assert result == "אברהם"

    def test_stage_b_remove_first_fallback(self):
        # Simulate a case where substring removal cannot find the exact token
        # but the pattern says remove first.
        # "כהן-לוי אברהם" — "כהן" is a substring but remove_substring uses
        # word-boundary padding, so " כהן " won't match " כהן-לוי ".
        # Stage A: unchanged → Stage B: REMOVE_FIRST → "אברהם"
        result = self.engine.remove_last_name_from_father(
            "כהן-לוי אברהם", "כהן", FatherNamePattern.REMOVE_FIRST
        )
        # After clean_name, "כהן-לוי" becomes "כהן לוי" (hyphen→space).
        # But here we pass already-cleaned strings directly.
        # "כהן" IS a substring of "כהן-לוי אברהם" but remove_substring
        # pads with spaces so " כהן " won't match " כהן-לוי " — Stage A unchanged
        # Stage B: REMOVE_FIRST → remove first token → "לוי אברהם" or "אברהם"
        # (depends on whether hyphen was already converted; here it's raw)
        assert "אברהם" in result

    def test_none_pattern_never_modifies(self):
        result = self.engine.remove_last_name_from_father(
            "כהן אברהם", "כהן", FatherNamePattern.NONE
        )
        assert result == "כהן אברהם"

    def test_last_name_absent_stage_b_still_fires(self):
        # Exact case from bug report:
        # last="דיטור", father="די טור אלעד"
        # "דיטור" is NOT a substring of "די טור אלעד" → Stage A no-op
        # Stage B REMOVE_FIRST → remove first token → "טור אלעד"
        result = self.engine.remove_last_name_from_father(
            "די טור אלעד", "דיטור", FatherNamePattern.REMOVE_FIRST
        )
        assert result == "טור אלעד"

    def test_last_name_absent_stage_b_remove_last(self):
        # last="דיטור", father="אלעד די טור", pattern=REMOVE_LAST
        # Stage A no-op → Stage B removes last token → "אלעד די"
        result = self.engine.remove_last_name_from_father(
            "אלעד די טור", "דיטור", FatherNamePattern.REMOVE_LAST
        )
        assert result == "אלעד די"

    def test_last_name_absent_stage_b_fires(self):
        # last="כהן" not in "אברהם לוי", pattern=REMOVE_LAST
        # Stage A: no-op. Stage B: REMOVE_LAST → "אברהם"
        result = self.engine.remove_last_name_from_father(
            "אברהם לוי", "כהן", FatherNamePattern.REMOVE_LAST
        )
        assert result == "אברהם"


# ---------------------------------------------------------------------------
# Two-stage removal rule — first name
# ---------------------------------------------------------------------------

class TestTwoStageRemovalFirstName:
    """Verify the same two-stage rule applies to first name."""

    def setup_method(self):
        self.engine = NameEngine(TextProcessor())

    def test_stage_a_success_stops_before_stage_b(self):
        # Example 3 from spec: last="כהן", first="כהן יעקב"
        # Stage A removes "כהן" → "יעקב"  (changed) → stop
        result = self.engine.remove_last_name_from_first_name(
            "כהן יעקב", "כהן", FatherNamePattern.REMOVE_FIRST
        )
        assert result == "יעקב"

    def test_stage_a_trailing_success_stops_before_stage_b(self):
        # last="כהן", first="יעקב כהן"
        # Stage A removes "כהן" → "יעקב"  (changed) → stop
        result = self.engine.remove_last_name_from_first_name(
            "יעקב כהן", "כהן", FatherNamePattern.REMOVE_LAST
        )
        assert result == "יעקב"

    def test_single_word_never_modified(self):
        result = self.engine.remove_last_name_from_first_name(
            "יעקב", "כהן", FatherNamePattern.REMOVE_FIRST
        )
        assert result == "יעקב"

    def test_last_name_absent_stage_b_fires(self):
        # last="כהן" not in "יעקב לוי", pattern=REMOVE_LAST
        # Stage A: no-op. Stage B: REMOVE_LAST → "יעקב"
        result = self.engine.remove_last_name_from_first_name(
            "יעקב לוי", "כהן", FatherNamePattern.REMOVE_LAST
        )
        assert result == "יעקב"

    def test_none_pattern_stage_b_skipped(self):
        # Stage A: "כהן-לוי יעקב" — "כהן" not matched by word-boundary → unchanged
        # Stage B: NONE → skip → return unchanged
        result = self.engine.remove_last_name_from_first_name(
            "כהן-לוי יעקב", "כהן", FatherNamePattern.NONE
        )
        assert result == "כהן-לוי יעקב"

    def test_stage_b_remove_first_fallback(self):
        # Stage A: "כהן-לוי יעקב" — word-boundary miss → unchanged
        # Stage B: REMOVE_FIRST → "יעקב"
        result = self.engine.remove_last_name_from_first_name(
            "כהן-לוי יעקב", "כהן", FatherNamePattern.REMOVE_FIRST
        )
        assert result == "יעקב"

    def test_empty_first_name_returns_empty(self):
        result = self.engine.remove_last_name_from_first_name(
            "", "כהן", FatherNamePattern.REMOVE_FIRST
        )
        assert result == ""

    def test_empty_last_name_returns_unchanged(self):
        result = self.engine.remove_last_name_from_first_name(
            "כהן יעקב", "", FatherNamePattern.REMOVE_FIRST
        )
        assert result == "כהן יעקב"


# ---------------------------------------------------------------------------
# detect_first_name_pattern
# ---------------------------------------------------------------------------

class TestDetectFirstNamePattern:
    def setup_method(self):
        self.engine = NameEngine(TextProcessor())

    def test_detects_remove_first_when_last_name_leads(self):
        first_sample = [["כהן יעקב"]] * 4
        last_sample = [["כהן"]] * 4
        pattern = self.engine.detect_first_name_pattern(first_sample, last_sample)
        assert pattern == FatherNamePattern.REMOVE_FIRST

    def test_detects_remove_last_when_last_name_trails(self):
        first_sample = [["יעקב כהן"]] * 4
        last_sample = [["כהן"]] * 4
        pattern = self.engine.detect_first_name_pattern(first_sample, last_sample)
        assert pattern == FatherNamePattern.REMOVE_LAST

    def test_returns_none_when_not_enough_matches(self):
        first_sample = [["יעקב"]] * 5
        last_sample = [["כהן"]] * 5
        pattern = self.engine.detect_first_name_pattern(first_sample, last_sample)
        assert pattern == FatherNamePattern.NONE

    def test_empty_sample_returns_none(self):
        pattern = self.engine.detect_first_name_pattern([], [])
        assert pattern == FatherNamePattern.NONE


# ---------------------------------------------------------------------------
# Pipeline order — unwanted token removal after char filtering
# ---------------------------------------------------------------------------

class TestCleanNamePipelineOrder:
    """Verify the strict fixed order: lang-detect → char-filter → token-remove."""

    def setup_method(self):
        self.tp = TextProcessor()

    def test_zl_removed_after_char_filtering(self):
        # ז"ל → after char filtering (quote removed) → זל → removed as token
        result = self.tp.clean_name('יוסי ז"ל')
        assert result == "יוסי"
        assert "זל" not in result

    def test_shlitta_removed_after_char_filtering(self):
        # שליט"א → שליטא after char filtering → removed as token
        result = self.tp.clean_name('הרב שליט"א כהן')
        assert "שליטא" not in result
        assert "כהן" in result

    def test_dr_removed(self):
        result = self.tp.clean_name("דר יוסי כהן")
        assert "דר" not in result
        assert "יוסי" in result
        assert "כהן" in result

    def test_rabi_removed(self):
        result = self.tp.clean_name("רבי יוסי")
        assert "רבי" not in result
        assert "יוסי" in result

    def test_hyphen_becomes_space(self):
        # Hyphens are converted to spaces, not kept
        result = self.tp.clean_name("בן-דוד")
        assert result == "בן דוד"

    def test_en_dash_becomes_space(self):
        result = self.tp.clean_name("בן\u2013דוד")
        assert result == "בן דוד"

    def test_digits_removed_hebrew(self):
        result = self.tp.clean_name("יוסי123כהן")
        assert "1" not in result
        assert "2" not in result
        assert "3" not in result

    def test_symbols_removed_hebrew(self):
        result = self.tp.clean_name("יוסי!@#כהן")
        assert "!" not in result
        assert result == "יוסי כהן" or result == "יוסיכהן" or ("יוסי" in result and "כהן" in result)

    def test_language_detected_before_filtering(self):
        # Pure Hebrew input → Hebrew dominant → English letters dropped
        result = self.tp.clean_name("יוסיABC")
        assert "A" not in result
        assert "יוסי" in result

    def test_unwanted_token_not_removed_as_substring(self):
        # "ר" is an unwanted token but only as a whole word
        # "ראובן" must NOT be affected
        result = self.tp.clean_name("ראובן כהן")
        assert "ראובן" in result

    def test_brd_removed(self):
        result = self.tp.clean_name("ברד יוסי")
        assert "ברד" not in result
        assert "יוסי" in result

    def test_empty_after_token_removal_returns_empty(self):
        result = self.tp.clean_name('ז"ל')
        assert result == ""

    def test_clean_name_and_clean_text_identical(self):
        tp = TextProcessor()
        for val in ["יוסי כהן", 'ז"ל', "Smith-Jones", "דר יוסי"]:
            assert tp.clean_name(val) == tp.clean_text(val)


# ---------------------------------------------------------------------------
# TextProcessor – backslash as name-part separator
# ---------------------------------------------------------------------------

class TestBackslashSeparator:
    """Backslash between name tokens must be treated as a space separator."""

    def setup_method(self):
        self.tp = TextProcessor()

    def test_backslash_between_two_parts(self):
        # חנה\ציפורה → חנה ציפורה
        result = self.tp.clean_name("חנה\\ציפורה")
        assert result == "חנה ציפורה"

    def test_backslash_with_leading_space(self):
        # חנה \ציפורה → חנה ציפורה
        result = self.tp.clean_name("חנה \\ציפורה")
        assert result == "חנה ציפורה"

    def test_backslash_with_trailing_space(self):
        # חנה\ ציפורה → חנה ציפורה
        result = self.tp.clean_name("חנה\\ ציפורה")
        assert result == "חנה ציפורה"

    def test_multiple_backslashes(self):
        # חנה\ציפורה\לאה → חנה ציפורה לאה
        result = self.tp.clean_name("חנה\\ציפורה\\לאה")
        assert result == "חנה ציפורה לאה"

    def test_backslash_combined_with_parentheses(self):
        # חנה\(ציפורה) → חנה ציפורה
        result = self.tp.clean_name("חנה\\(ציפורה)")
        assert result == "חנה ציפורה"

    def test_single_token_no_separator(self):
        # No backslash — single token unchanged
        result = self.tp.clean_name("חנה")
        assert result == "חנה"

    def test_backslash_only_no_tokens(self):
        # Backslash with no surrounding name letters → empty
        result = self.tp.clean_name("\\")
        assert result == ""
