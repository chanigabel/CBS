"""derived_columns: inject serial-number and MosadID derived columns.

Both the UI (WorkbookService.get_sheet_data) and the export pipeline
(ExportService.visible_rows) call these helpers so the two paths are
always consistent.

Serial number
-------------
The source workbook may contain a serial-number column under various
Hebrew/English header names (e.g. "מספר סידורי", "Serial No.", "מס'",
"#", …).  The extractor passes it through as a sanitised key.

We detect it by scanning field_names for any key whose *original header
text* (stored in the passthrough key itself, after un-sanitising) matches
known serial-number patterns.  If found and the row already has a value
we keep it; if the cell is blank we auto-fill with the 1-based visible
row position.  If no serial-number column exists in the source at all we
inject a synthetic ``_serial`` column at position 0.

MosadID
-------
Already injected into row dicts by WorkbookService (from sheet metadata).
Here we ensure it appears in display_columns immediately after the serial
number column, whether it came from a real source column or from metadata.
"""

import re
import unicodedata
from typing import Any, Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# Serial-number detection
# ---------------------------------------------------------------------------

# Normalised fragments that identify a serial-number column header.
_SERIAL_FRAGMENTS = [
    "מספר סידורי",
    "מס סידורי",
    "מס' סידורי",
    "מספר שורה",
    "serial no",
    "serial number",
    "serial",
    "row no",
    "row number",
    "#",
    "מס'",
    "מספר",   # generic "number" — only matched when it's the *entire* key
]

# Fragments that must match the *entire* normalised key (not just a substring)
# to avoid false positives on e.g. "id_number".
_SERIAL_EXACT_ONLY = {"מספר", "#", "מס'"}


def _norm(s: str) -> str:
    """NFC-normalise, lower-case, collapse whitespace, strip punctuation."""
    s = unicodedata.normalize("NFC", s)
    s = s.replace("_", " ").replace("-", " ")
    s = re.sub(r"[^\w\u0590-\u05FF\s]", " ", s)
    s = " ".join(s.split()).lower()
    return s


def _key_to_readable(key: str) -> str:
    """Convert a sanitised field key back to a human-readable form."""
    return key.replace("_", " ").strip()


def detect_serial_field(field_names: List[str]) -> Optional[str]:
    """Return the field_name that is the serial-number column, or None.

    Scans field_names in order and returns the first match.
    """
    for fname in field_names:
        readable = _norm(_key_to_readable(fname))
        for frag in _SERIAL_FRAGMENTS:
            nfrag = _norm(frag)
            if frag in _SERIAL_EXACT_ONLY:
                if readable == nfrag:
                    return fname
            else:
                if nfrag in readable:
                    return fname
    return None


# Internal synthetic key used when no serial column exists in the source.
SYNTHETIC_SERIAL_KEY = "_serial"
MOSAD_ID_KEY = "MosadID"


def apply_derived_columns(
    rows: List[Dict[str, Any]],
    field_names: List[str],
    display_columns: List[str],
    meta_mosad_id: Optional[str] = None,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Inject serial-number and MosadID derived columns.

    Mutates *rows* in-place (adds/fills keys) and returns a new
    display_columns list with the two derived columns placed at the front
    (serial first, MosadID second).

    Args:
        rows:            Visible rows (already filtered, metadata-stripped).
        field_names:     Original source field names from SheetDataset.
        display_columns: Current display column list built by WorkbookService.
        meta_mosad_id:   MosadID from sheet metadata (may be None).

    Returns:
        (rows, new_display_columns)
    """
    # ------------------------------------------------------------------
    # 1. Serial number
    # ------------------------------------------------------------------
    serial_field = detect_serial_field(field_names)

    if serial_field is None:
        # No serial column in source — inject synthetic one.
        serial_col = SYNTHETIC_SERIAL_KEY
        for i, row in enumerate(rows, start=1):
            row[SYNTHETIC_SERIAL_KEY] = i
    else:
        serial_col = serial_field
        # Fill blanks with auto-generated position.
        for i, row in enumerate(rows, start=1):
            v = row.get(serial_field)
            if v is None or str(v).strip() == "":
                row[serial_field] = i

    # ------------------------------------------------------------------
    # 2. MosadID — inject from metadata into rows that don't have it.
    # ------------------------------------------------------------------
    if meta_mosad_id is not None:
        for row in rows:
            if not row.get(MOSAD_ID_KEY):
                row[MOSAD_ID_KEY] = meta_mosad_id

    # ------------------------------------------------------------------
    # 3. Build new display_columns: serial first, MosadID second, rest after.
    # ------------------------------------------------------------------
    # Remove serial and MosadID from wherever they currently sit.
    rest = [
        c for c in display_columns
        if c != serial_col and c != MOSAD_ID_KEY and c != SYNTHETIC_SERIAL_KEY
    ]

    # Decide whether MosadID column should appear:
    # show it if any row has a non-empty value for it.
    mosad_id_has_value = any(
        row.get(MOSAD_ID_KEY) not in (None, "")
        for row in rows
    )

    new_display = [serial_col]
    if mosad_id_has_value:
        new_display.append(MOSAD_ID_KEY)
    new_display.extend(rest)

    return rows, new_display
