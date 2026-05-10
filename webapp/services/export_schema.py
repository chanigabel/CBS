"""Canonical export schema and field mapping helpers."""

from __future__ import annotations

import unicodedata
from typing import Dict, List, Optional

# ---------------------------------------------------------------------------
# Sheet-name normalization and canonical export name mapping.
# ---------------------------------------------------------------------------

_SHEET_NAME_PATTERNS = [
    ("DayarimYahidim", ["דיירים יחידים", "דיירים"]),
    ("MeshkeyBayt", ["מתגוררים במשקי בית", "משקי בית", "מתגוררים"]),
    ("AnasheyTzevet", ["אנשי צוות ובני משפחותיהם", "אנשי צוות", "צוות"]),
]


# מנרמל שם גיליון לפני התאמתו לסכמת היצוא.
def _normalize_text(text: str) -> str:
    """Strip, collapse whitespace, and apply Unicode NFC normalization."""
    return unicodedata.normalize("NFC", " ".join(text.split()))


# מחזיר שם גיליון קנוני שמכתיב מיפוי וכותרות יצוא.
def canonical_sheet_name(source_name: str) -> str:
    """Map a source sheet name to its canonical export name."""
    normalized = _normalize_text(source_name)
    for export_name, keywords in _SHEET_NAME_PATTERNS:
        for keyword in keywords:
            if _normalize_text(keyword) in normalized:
                return export_name
    return source_name


# ---------------------------------------------------------------------------
# Per-sheet-type export schemas (column order matters).
# ---------------------------------------------------------------------------

_HEADERS_DAYARIM: List[str] = [
    "MosadID", "SugMosad",
    "ShemPrati", "ShemMishpaha", "ShemHaAv",
    "MisparZehut", "Darkon", "Min",
    "ShnatLida", "HodeshLida", "YomLida",
    "shnatknisa", "Hodeshknisa", "YomKnisa",
]

_HEADERS_MESHKEY: List[str] = [
    "MosadID", "SugMosad", "MisparDiraBeMosad",
    "ShemPrati", "ShemMishpaha", "ShemHaAv",
    "MisparZehut", "Darkon", "Min",
    "ShnatLida", "HodeshLida", "YomLida",
    "ShnatKnisa", "HodeshKnisa", "YomKnisa",
]

_HEADERS_DEFAULT = _HEADERS_DAYARIM

_SCHEMA_BY_CANONICAL: Dict[str, List[str]] = {
    "DayarimYahidim": _HEADERS_DAYARIM,
    "MeshkeyBayt": _HEADERS_MESHKEY,
    "AnasheyTzevet": _HEADERS_MESHKEY,
}


# מחזיר את כותרות היצוא התקניות לפי סוג גיליון.
def headers_for_sheet(canonical_name: str) -> List[str]:
    """Return the ordered column list for the given canonical sheet name."""
    return _SCHEMA_BY_CANONICAL.get(canonical_name, _HEADERS_DEFAULT)


# ---------------------------------------------------------------------------
# Field mapping: export header → source JSON key.
# ---------------------------------------------------------------------------

EXPORT_MAPPING: Dict[str, Optional[str]] = {
    "MosadID": "MosadID",
    "SugMosad": "SugMosad",
    "MisparDiraBeMosad": "MisparDiraBeMosad",
    "ShemPrati": "first_name_corrected",
    "ShemMishpaha": "last_name_corrected",
    "ShemHaAv": "father_name_corrected",
    "MisparZehut": "id_number_corrected",
    "Darkon": "passport_corrected",
    "Min": "gender_corrected",
    "ShnatLida": "birth_year_corrected",
    "HodeshLida": "birth_month_corrected",
    "YomLida": "birth_day_corrected",
    "shnatknisa": "entry_year_corrected",
    "ShnatKnisa": "entry_year_corrected",
    "Hodeshknisa": "entry_month_corrected",
    "HodeshKnisa": "entry_month_corrected",
    "YomKnisa": "entry_day_corrected",
}
