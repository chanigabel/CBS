"""sheet_name_resolver: map source sheet names to canonical institution-report names.

Canonical names:
    AnasheyTzevet   — "אנשי צוות ובני משפחותיהם" and variants
    DayarimYahidim  — "דיירים יחידים" and variants
    MeshkeyBayt     — "מתגוררים במשקי בית" and variants

This module is intentionally thin — it mirrors the logic already present in
webapp/services/export_service.py (canonical_sheet_name) so that the
validation layer can resolve sheet names without importing from the webapp.
"""

import unicodedata

# Canonical name → list of keyword fragments (any must match after NFC normalisation)
_SHEET_NAME_PATTERNS = [
    ("DayarimYahidim",  ["דיירים יחידים", "דיירים"]),
    ("MeshkeyBayt",     ["מתגוררים במשקי בית", "משקי בית", "מתגוררים"]),
    ("AnasheyTzevet",   ["אנשי צוות ובני משפחותיהם", "אנשי צוות", "צוות"]),
]


# מנרמל שם גיליון כדי לזהות אותו גם כשיש רווחים או סימנים שונים.
def _normalize_text(s: str) -> str:
    return unicodedata.normalize("NFC", " ".join(s.split()))


# ממפה שם גיליון מהקובץ לשם קנוני המשמש validation ו־export.
def resolve_canonical_sheet_name(source_name: str) -> str:
    """Map a source sheet name to its canonical institution-report name.

    Returns the canonical name if matched, or the original name unchanged.
    Matching is case-insensitive and NFC-normalised.

    Examples:
        "דיירים יחידים"                    → "DayarimYahidim"
        "אנשי צוות ובני משפחותיהם"         → "AnasheyTzevet"
        "מתגוררים במשקי בית"               → "MeshkeyBayt"
        "DayarimYahidim"                   → "DayarimYahidim"  (already canonical)
        "SomeOtherSheet"                   → "SomeOtherSheet"  (unchanged)
    """
    normalised = _normalize_text(source_name)
    # Also accept already-canonical English names directly.
    for export_name, keywords in _SHEET_NAME_PATTERNS:
        if normalised == export_name:
            return export_name
        for kw in keywords:
            if _normalize_text(kw) in normalised:
                return export_name
    return source_name
