"""Low-level text normalization helpers for name fields.

The processor converts arbitrary cell values into cleaned display strings
used by the name engine. It removes invisible characters, normalizes spacing,
strips diacritics, filters to the dominant script, and removes configured
titles or tokens.
"""

from ..data_types import Language
import re

# Matches a parenthesized group whose content contains at least one Hebrew
# quote/acronym character (" ״ ׳ ').  The entire group (including parens) is
# removed by clean_name before character filtering runs.
_RE_PAREN_ACRONYM = re.compile(r'\([^)]*["\u05f4\u05f3\'"][^)]*\)')


class TextProcessor:
    """Pure business logic for text manipulation."""

    # Hebrew letter Unicode range: 1488-1514
    HEBREW_START = 0x05D0  # 1488
    HEBREW_END = 0x05EA  # 1514

    # Hebrew final letters: ך, ם, ן, ף, ץ
    HEBREW_FINAL_LETTERS = {
        "\u05da",  # ך (Final Kaf)
        "\u05dd",  # ם (Final Mem)
        "\u05df",  # ן (Final Nun)
        "\u05e3",  # ף (Final Pe)
        "\u05e5",  # ץ (Final Tsadi)
    }

    # Valid separators — kept for backwards-compat with code that reads this set
    VALID_SEPARATORS = {" ", "-", "\u2013", "\u2014"}

    # Diacritic mappings (character to base character)
    DIACRITIC_MAP = {
        "à": "a", "á": "a", "â": "a", "ã": "a", "ä": "a", "å": "a",
        "è": "e", "é": "e", "ê": "e", "ë": "e",
        "ì": "i", "í": "i", "î": "i", "ï": "i",
        "ò": "o", "ó": "o", "ô": "o", "õ": "o", "ö": "o",
        "ù": "u", "ú": "u", "û": "u", "ü": "u",
        "ý": "y", "ÿ": "y", "ñ": "n", "ç": "c",
        "À": "A", "Á": "A", "Â": "A", "Ã": "A", "Ä": "A", "Å": "A",
        "È": "E", "É": "E", "Ê": "E", "Ë": "E",
        "Ì": "I", "Í": "I", "Î": "I", "Ï": "I",
        "Ò": "O", "Ó": "O", "Ô": "O", "Õ": "O", "Ö": "O",
        "Ù": "U", "Ú": "U", "Û": "U", "Ü": "U",
        "Ý": "Y", "Ñ": "N", "Ç": "C",
        "\u0451": "e",  # Cyrillic ё
    }

    # Hebrew honorific titles — raw form (still have punctuation).
    # Kept for backwards-compat with code that calls remove_titles() directly.
    HEBREW_TITLES = [
        "ז\"ל",
        "זצ\"ל",
        "זיע\"א",
        "הי\"ד",
        "שליט\"א",
    ]

    # English honorific titles — raw form.
    ENGLISH_TITLES = [
        "mr.", "mrs.", "ms.", "dr.", "prof.", "jr.", "sr.", "iii", "iv",
    ]

    # Unwanted Hebrew tokens matched AFTER character filtering.
    # Punctuation has been removed by then, so ז"ל → זל, שליט"א → שליטא, etc.
    HEBREW_UNWANTED_TOKENS = {
        "זל",       # ז"ל after cleanup
        "זצל",      # זצ"ל after cleanup
        "זיעא",     # זיע"א after cleanup
        "היד",      # הי"ד after cleanup
        "שליטא",    # שליט"א after cleanup
        "דר",       # ד"ר / doctor
        "רבי",      # rabbi title
        "ר",        # abbreviated rabbi (whole-token only)
        "ברד",
        "ברמ",
        "בראא",
        "בראש",
        "בימ",
        "ברדא",
        "ברי",
    }

    # Subset of HEBREW_UNWANTED_TOKENS that are name-prefix titles.
    # These words CAN stand alone as a valid name/title (e.g. a person
    # known only as "רבי"), so they must NOT be removed when they are the
    # sole remaining token in the field.  All other unwanted tokens (memorial
    # honorifics, abbreviations) are always removed regardless of word count.
    _NAME_PREFIX_TITLES = {
        "רבי",   # rabbi — may be a standalone name
        "ר",     # abbreviated rabbi — may be a standalone name
    }

    # All hyphen-like characters — converted to spaces during char filtering
    _HYPHEN_CHARS = {
        "-",        # ASCII hyphen-minus
        "\u2010",   # hyphen
        "\u2011",   # non-breaking hyphen
        "\u2012",   # figure dash
        "\u2013",   # en-dash
        "\u2014",   # em-dash
        "\u2015",   # horizontal bar
        "\u2212",   # minus sign
    }

    # Zero-width / invisible Unicode characters stripped at the start
    _ZERO_WIDTH = {
        "\u200b", "\u200c", "\u200d", "\u200e", "\u200f",
        "\u202a", "\u202b", "\u202c", "\u202d", "\u202e", "\ufeff",
    }

    # Arabic-Indic digit translation table (built once)
    _ARABIC_INDIC = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")

    # ------------------------------------------------------------------
    # Low-level helpers
    # ------------------------------------------------------------------

    def safe_to_string(self, value) -> str:
        """Safely convert any variant to string."""
        if value is None:
            return ""
        try:
            return str(value)
        except Exception:
            return ""

    def minimal_normalize(self, text: str) -> str:
        """Trim + collapse spaces + strip zero-width chars."""
        text = "".join(ch for ch in text if ch not in self._ZERO_WIDTH)
        return " ".join(text.strip().split())

    def worksheet_trim(self, text: str) -> str:
        """WorksheetFunction.Trim equivalent: trim + collapse internal spaces."""
        return " ".join(self.safe_to_string(text).split())

    def collapse_spaces(self, text: str) -> str:
        """Replace multiple consecutive spaces with a single space."""
        return " ".join(self.safe_to_string(text).split())

    # ------------------------------------------------------------------
    # Title / substring removal (kept for backwards-compat)
    # ------------------------------------------------------------------

    def remove_titles(self, text: str) -> str:
        """Remove raw-form Hebrew/English titles (before char filtering).

        Kept for backwards-compat with code that calls this directly.
        In the main clean_name pipeline, unwanted-token removal now happens
        AFTER char filtering via remove_unwanted_tokens().
        """
        if not text:
            return ""

        padded = f" {text} "

        for title in self.HEBREW_TITLES:
            if title in padded:
                padded = padded.replace(title, " ")

        lower_padded = padded.lower()
        for title in self.ENGLISH_TITLES:
            t = f" {title} "
            if t in lower_padded:
                idx = lower_padded.find(t)
                while idx != -1:
                    padded = padded[:idx] + " " + padded[idx + len(t):]
                    lower_padded = lower_padded[:idx] + " " + lower_padded[idx + len(t):]
                    idx = lower_padded.find(t)

        return self.worksheet_trim(padded)

    def remove_unwanted_tokens(self, text: str) -> str:
        """Remove unwanted Hebrew tokens from already-cleaned text.

        Must be called AFTER character filtering so that punctuation has been
        removed (e.g. ז"ל → זל before this runs).

        Tokens are matched as whole words using space-padded boundaries.

        Single-word preservation rule: if the entire value consists of only
        one word and that word is a name-prefix title (e.g. "רבי"), the word
        is kept as-is.  This applies only to tokens in _NAME_PREFIX_TITLES —
        memorial honorifics and abbreviations are always removed regardless
        of word count.
        """
        if not text:
            return ""

        # Single-word guard: preserve name-prefix titles when they are the
        # only word in the field (e.g. "רבי" alone is a valid standalone name).
        words = text.split()
        if len(words) == 1 and words[0] in self._NAME_PREFIX_TITLES:
            return text

        padded = f" {text} "
        for token in self.HEBREW_UNWANTED_TOKENS:
            padded = padded.replace(f" {token} ", " ")

        return self.worksheet_trim(padded)

    def remove_substring(self, text: str, substring: str) -> str:
        """Remove a word/phrase from text (word-boundary aware, VBA parity)."""
        base = self.safe_to_string(text)
        sub = self.safe_to_string(substring)
        if not base or not sub:
            return self.worksheet_trim(base)

        padded_text = f" {base} "
        padded_sub = f" {sub} "
        result = padded_text.replace(padded_sub, " ")
        return self.worksheet_trim(result)

    # ------------------------------------------------------------------
    # Diacritics, language detection, final-letter spacing
    # ------------------------------------------------------------------

    def remove_diacritics(self, text: str) -> str:
        """Remove diacritics using the DIACRITIC_MAP."""
        return "".join(self.DIACRITIC_MAP.get(ch, ch) for ch in text)

    def detect_language_dominance(self, text: str) -> Language:
        """Detect dominant language by counting Hebrew vs English letters.

        Hebrew wins on tie.
        """
        hebrew_count = 0
        english_count = 0

        for ch in text:
            code = ord(ch)
            if self.HEBREW_START <= code <= self.HEBREW_END:
                hebrew_count += 1
            elif ("A" <= ch <= "Z") or ("a" <= ch <= "z"):
                english_count += 1

        if hebrew_count == 0 and english_count == 0:
            return Language.MIXED

        if hebrew_count >= english_count:
            return Language.HEBREW
        return Language.ENGLISH

    def fix_hebrew_final_letters(self, text: str) -> str:
        """Insert a space after final Hebrew letters when followed by a non-space char."""
        if not text:
            return ""

        result_chars = []
        for i, ch in enumerate(text):
            result_chars.append(ch)
            if ch in self.HEBREW_FINAL_LETTERS and i + 1 < len(text):
                next_ch = text[i + 1]
                if next_ch not in {" ", ",", ".", ";", ":", "!", "?", "-", "\u2013", "\u2014"}:
                    result_chars.append(" ")

        return "".join(result_chars)

    # ------------------------------------------------------------------
    # Public entry point — strict fixed-order pipeline
    # ------------------------------------------------------------------

    def clean_name(self, value) -> str:
        """Clean a name value using a strict fixed-order pipeline.

        Order of operations:
            1. SafeToString + strip zero-width characters
            2. Diacritic removal (so accented Latin letters count correctly)
            3. Language detection — count Hebrew vs English letters only
            4. Character filtering:
               - Keep only dominant-language letters
               - Convert all hyphen-like characters to spaces
               - Drop everything else (digits, symbols, wrong-language letters)
            5. Space normalisation — trim + collapse multiple spaces
            6. Unwanted token removal — on the cleaned form, so ז"ל → זל
               is matched correctly after punctuation has been removed
        """
        # 1. SafeToString + strip zero-width
        text = self.safe_to_string(value)
        if not text:
            return ""
        text = "".join(ch for ch in text if ch not in self._ZERO_WIDTH)
        if not text:
            return ""

        # 2. Diacritic removal + Arabic-Indic digit normalisation
        text = self.remove_diacritics(text)
        text = text.translate(self._ARABIC_INDIC)

        # 3. Language detection
        language = self.detect_language_dominance(text)

        # 3b. Remove parenthesized acronym tokens BEFORE character filtering.
        # If the content inside parentheses contains a Hebrew quote/acronym
        # character (" ״ ׳ '), the entire parenthesized group is discarded.
        # Normal words in parentheses (no quote chars) are kept — the parens
        # themselves will be converted to spaces in step 4 as usual.
        text = _RE_PAREN_ACRONYM.sub("", text)

        # 4. Character filtering
        filtered: list = []
        for ch in text:
            code = ord(ch)
            is_hebrew = self.HEBREW_START <= code <= self.HEBREW_END
            is_english = ("A" <= ch <= "Z") or ("a" <= ch <= "z")

            if ch == " ":
                filtered.append(" ")
            elif ch in self._HYPHEN_CHARS:
                filtered.append(" ")          # hyphens → space
            elif ch in ("(", ")"):
                filtered.append(" ")          # parentheses → space (so adjacent tokens are separated)
            elif ch == "\\":
                filtered.append(" ")          # backslash → space (so adjacent tokens are separated)
            elif language == Language.HEBREW and is_hebrew:
                filtered.append(ch)
            elif language == Language.ENGLISH and is_english:
                filtered.append(ch)
            elif language == Language.MIXED and (is_hebrew or is_english):
                filtered.append(ch)
            # Everything else dropped

        text = "".join(filtered)

        # 5. Space normalisation
        text = " ".join(text.split())
        if not text:
            return ""

        # 6. Unwanted token removal (on cleaned form)
        if language in (Language.HEBREW, Language.MIXED):
            text = self.remove_unwanted_tokens(text)
        elif language == Language.ENGLISH:
            # English titles have lost their trailing dot after char filtering.
            # Match case-insensitively but preserve the original casing in output.
            padded = f" {text} "
            lower_padded = f" {text.lower()} "
            for title in self.ENGLISH_TITLES:
                clean_title = title.rstrip(".")
                target = f" {clean_title} "
                idx = lower_padded.find(target)
                while idx != -1:
                    padded = padded[:idx] + " " + padded[idx + len(target):]
                    lower_padded = lower_padded[:idx] + " " + lower_padded[idx + len(target):]
                    idx = lower_padded.find(target)
            text = self.worksheet_trim(padded)

        return text

    def clean_text(self, text: str) -> str:
        """Legacy alias for clean_name."""
        return self.clean_name(text)
