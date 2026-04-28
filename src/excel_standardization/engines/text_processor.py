"""TextProcessor — shared text cleaning utilities for name standardization.

Purpose:
    Provides the low-level text pipeline used by NameEngine to clean name
    field values.  All methods are pure Python with no Excel or web
    dependencies.  The main entry point is ``clean_name()``.

Implemented rules (in pipeline order inside ``clean_name``):
    1. Safe string conversion
       Converts any input type (None, int, float, etc.) to a string safely.
       None → "".
       Example:
           None        → ""
           123         → "123"  (then digits are dropped in step 4)

    2. Zero-width / invisible character removal
       Strips Unicode zero-width and directional control characters before
       any other processing.
       Characters removed: U+200B, U+200C, U+200D, U+200E, U+200F,
                           U+202A–U+202E, U+FEFF.
       Example:
           "יוסי\u200bכהן"  → "יוסיכהן"  (then space-normalised)

    3. Diacritic removal
       Replaces accented Latin characters with their ASCII base equivalents
       using a fixed mapping of ~40 characters.
       Example:
           "José"  → "Jose"
           "Müller" → "Muller"
       Also translates Arabic-Indic digits (٠١٢٣٤٥٦٧٨٩) to ASCII digits
       (0–9) before the digit-drop step.

    4. Language detection
       Counts Hebrew letters (U+05D0–U+05EA) vs. ASCII letters (A–Z, a–z).
       Hebrew wins on a tie.  Result is one of: HEBREW, ENGLISH, MIXED.
       MIXED is returned only when both counts are zero (no letters at all).
       Example:
           "יוסי כהן"   → HEBREW
           "John Smith" → ENGLISH
           "123 !!!"    → MIXED  (no letters)

    5. Character filtering (dominant language kept, everything else dropped)
       Rules applied per character:
       - Spaces → kept as-is.
       - Hyphen-like characters (ASCII hyphen, en-dash, em-dash, minus, etc.)
         → converted to a single space.
       - Parentheses ( ) → converted to a space (so adjacent tokens separate).
       - Backslash \\ → converted to a space.
       - Hebrew letters → kept only when language is HEBREW or MIXED.
       - ASCII letters → kept only when language is ENGLISH or MIXED.
       - Digits → always dropped (regardless of language).
       - All other characters (symbols, punctuation, wrong-language letters)
         → dropped.
       Examples:
           "אבר9הם"        → "אברהם"   (digit dropped)
           "כהן-לוי"       → "כהן לוי" (hyphen → space)
           "Smith (Jr)"    → "Smith Jr" (parentheses → spaces)
           "John\\Jane"    → "John Jane" (backslash → space)
           "José123"       → "Jose"     (digits dropped, diacritics removed)

    6. Space normalisation
       Trims leading/trailing whitespace and collapses multiple consecutive
       spaces to a single space.
       Example:
           "  משה   חיים  "  → "משה חיים"

    7. Unwanted token removal (Hebrew and English)
       For Hebrew/MIXED text: removes a fixed set of tokens matched as whole
       words after character filtering.  Tokens include memorial honorifics
       (ז"ל → זל after filtering), religious titles, and abbreviations.
       Single-word preservation rule: if the entire value is a single word
       that is a name-prefix title ("רבי" or "ר"), it is kept as-is.
       For English text: removes English honorific titles (mr, mrs, ms, dr,
       prof, jr, sr, iii, iv) matched case-insensitively as whole words.
       Examples:
           'יוסי ז"ל'      → "יוסי"   (ז"ל → זל → removed)
           'ד"ר כהן'       → "כהן"    (ד"ר → דר → removed)
           "רבי"           → "רבי"    (single-word title preserved)
           "Dr. John"      → "John"   (English title removed)

Other public methods:
    safe_to_string(value)
        Converts any value to string; returns "" for None or on error.

    worksheet_trim(text)
        Equivalent to Excel WorksheetFunction.Trim: trims + collapses spaces.

    collapse_spaces(text)
        Collapses multiple consecutive spaces to one.

    minimal_normalize(text)
        Strips zero-width chars + trims + collapses spaces.  Does NOT do
        language detection or character filtering.

    remove_diacritics(text)
        Applies DIACRITIC_MAP only.  Does not do any other cleaning.

    detect_language_dominance(text)
        Returns Language.HEBREW, Language.ENGLISH, or Language.MIXED.

    fix_hebrew_final_letters(text)
        Inserts a space after a Hebrew final letter (ך ם ן ף ץ) when it is
        immediately followed by a non-space character.
        Example: "כהןלוי" → "כהן לוי"

    remove_substring(text, substring)
        Removes a word/phrase from text using space-padded word boundaries.
        Example: remove_substring("אברהם כהן", "כהן") → "אברהם"

    remove_unwanted_tokens(text)
        Removes HEBREW_UNWANTED_TOKENS as whole words from already-cleaned text.
        Must be called AFTER character filtering (punctuation already removed).

    remove_titles(text)
        Legacy method: removes raw-form Hebrew/English titles before character
        filtering.  Kept for backwards compatibility; the main pipeline uses
        remove_unwanted_tokens() instead.

Important notes:
    - This class does not read or write Excel files.
    - It does not call any web or I/O layer.
    - All methods are deterministic: same input always produces same output.
    - clean_name() is the only method that runs the full pipeline.
      All other methods are partial steps or helpers.

Known limitations:
    - Language detection is binary (Hebrew vs. English); mixed-script names
      (e.g. a Hebrew name with an English middle name) are handled by the
      MIXED path, which keeps both Hebrew and English letters.
    - The diacritic map covers common Western European characters but is not
      exhaustive (e.g. some Eastern European or Cyrillic characters beyond ё
      are not mapped).
    - Digit removal is unconditional: any digit in a name field is dropped,
      even if it is part of a legitimate name suffix (e.g. "John III" →
      "John" because "III" is removed as an English title, but "John 3rd"
      → "John rd" → "John" after token cleanup is not guaranteed).
"""

from ..data_types import Language


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
