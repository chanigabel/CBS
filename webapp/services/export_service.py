"""ExportService: writes the final normalized workbook to disk for download."""

import logging
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from fastapi import HTTPException
from openpyxl import Workbook
from openpyxl.styles import Alignment

from webapp.services.session_service import SessionService
from webapp.services.derived_columns import apply_derived_columns, detect_serial_field, SYNTHETIC_SERIAL_KEY

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Sheet-name normalisation and canonical export name mapping.
# ---------------------------------------------------------------------------

# Keywords that identify each sheet type (after Unicode normalisation + strip).
# Each entry: (canonical_export_name, list_of_keyword_fragments_any_must_match)
_SHEET_NAME_PATTERNS = [
    ("DayarimYahidim",  ["דיירים יחידים", "דיירים"]),
    ("MeshkeyBayt",     ["מתגוררים במשקי בית", "משקי בית", "מתגוררים"]),
    ("AnasheyTzevet",   ["אנשי צוות ובני משפחותיהם", "אנשי צוות", "צוות"]),
]


def _normalize_text(s: str) -> str:
    """Strip, collapse whitespace, and apply Unicode NFC normalisation."""
    return unicodedata.normalize("NFC", " ".join(s.split()))


def canonical_sheet_name(source_name: str) -> str:
    """Map a source sheet name to its canonical export name.

    Matching is done after NFC normalisation and whitespace collapsing.
    Returns the original name unchanged when no pattern matches.
    """
    normalised = _normalize_text(source_name)
    for export_name, keywords in _SHEET_NAME_PATTERNS:
        for kw in keywords:
            if _normalize_text(kw) in normalised:
                return export_name
    return source_name


# ---------------------------------------------------------------------------
# Per-sheet-type export schemas (column order matters).
# ---------------------------------------------------------------------------

# DayarimYahidim — 14 columns (no MisparDiraBeMosad)
_HEADERS_DAYARIM: List[str] = [
    "MosadID", "SugMosad",
    "ShemPrati", "ShemMishpaha", "ShemHaAv",
    "MisparZehut", "Darkon", "Min",
    "ShnatLida", "HodeshLida", "YomLida",
    "shnatknisa", "Hodeshknisa", "YomKnisa",
]

# MeshkeyBayt / AnasheyTzevet — 15 columns (includes MisparDiraBeMosad)
_HEADERS_MESHKEY: List[str] = [
    "MosadID", "SugMosad", "MisparDiraBeMosad",
    "ShemPrati", "ShemMishpaha", "ShemHaAv",
    "MisparZehut", "Darkon", "Min",
    "ShnatLida", "HodeshLida", "YomLida",
    "ShnatKnisa", "HodeshKnisa", "YomKnisa",
]

# Unknown / unmatched sheets fall back to the DayarimYahidim schema.
_HEADERS_DEFAULT = _HEADERS_DAYARIM

_SCHEMA_BY_CANONICAL: Dict[str, List[str]] = {
    "DayarimYahidim": _HEADERS_DAYARIM,
    "MeshkeyBayt":    _HEADERS_MESHKEY,
    "AnasheyTzevet":  _HEADERS_MESHKEY,
}


def headers_for_sheet(canonical_name: str) -> List[str]:
    """Return the ordered column list for the given canonical sheet name."""
    return _SCHEMA_BY_CANONICAL.get(canonical_name, _HEADERS_DEFAULT)


# ---------------------------------------------------------------------------
# Field mapping: export header → source JSON key.
# ---------------------------------------------------------------------------
# Rules:
#   - Corrected fields: use *_corrected keys only, no fallback to originals.
#   - MosadID / SugMosad / MisparDiraBeMosad: read from source row as-is;
#     if the key is absent or blank the cell is left empty.

EXPORT_MAPPING: Dict[str, Optional[str]] = {
    "MosadID":             "MosadID",
    "SugMosad":            "SugMosad",
    "MisparDiraBeMosad":   "MisparDiraBeMosad",
    "ShemPrati":           "first_name_corrected",
    "ShemMishpaha":        "last_name_corrected",
    "ShemHaAv":            "father_name_corrected",
    "MisparZehut":         "id_number_corrected",
    "Darkon":              "passport_corrected",
    "Min":                 "gender_corrected",
    "ShnatLida":           "birth_year_corrected",
    "HodeshLida":          "birth_month_corrected",
    "YomLida":             "birth_day_corrected",
    "ShnatKnisa":          "entry_year_corrected",
    "HodeshKnisa":         "entry_month_corrected",
    "YomKnisa":            "entry_day_corrected",
}


def _to_pascal_case(text: str) -> str:
    """Convert a free-text name to PascalCase English words joined without spaces.

    Examples:
        "Beit Haharon"  -> "BeitHaharon"
        "beit ha-baron" -> "BeitHaBaron"
        "מוסד הטוב"     -> "מוסד הטוב"  (non-ASCII kept as-is, joined)
    """
    import re
    # Split on whitespace and hyphens, capitalise each token, join
    tokens = re.split(r"[\s\-]+", text.strip())
    return "".join(t.capitalize() for t in tokens if t)


def _build_export_filename(record) -> str:
    """Build the export filename from institution metadata.

    Format: ``{MosadID} {MosadNamePascal}.xlsx``
    Falls back to the original stem + timestamp when fields are empty.
    """
    mosad_id = (record.mosad_id or "").strip()
    mosad_name = (record.mosad_name or "").strip()

    if mosad_id and mosad_name:
        pascal = _to_pascal_case(mosad_name)
        return f"{mosad_id} {pascal}.xlsx"

    # Fallback: original filename stem + timestamp
    original_stem = Path(record.original_filename).stem
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{original_stem}_normalized_{timestamp}.xlsx"


def _is_numeric_like(value: Any) -> bool:
    """Return True if value is numeric or a string that parses as a number."""
    if isinstance(value, (int, float)):
        return True
    try:
        float(str(value).strip())
        return True
    except (ValueError, TypeError):
        return False


def visible_rows(sheet_dataset) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Return (rows, display_columns) exactly as the UI would show them.

    Applies the same filters and derived-column logic that
    WorkbookService.get_sheet_data uses:

    1. Strip internal metadata keys (``_normalization*``).
    2. Drop completely empty rows (checked against original source columns).
    3. Drop the leading numbers-only helper row if present.
    4. Inject serial number and MosadID derived columns via apply_derived_columns.

    Returns a tuple of (filtered_rows, display_columns) so the export loop
    can use the same column list the UI shows.
    """
    original_field_set = set(sheet_dataset.field_names)

    # Strip internal metadata keys.
    rows = [
        {k: v for k, v in row.items() if not k.startswith("_normalization")}
        for row in sheet_dataset.rows
    ]

    # Drop completely empty rows (checked against original source columns only).
    rows = [
        row for row in rows
        if any(
            v is not None and str(v).strip() != ""
            for k, v in row.items()
            if k in original_field_set
        )
    ]

    # Drop leading numbers-only helper row if present.
    if rows:
        first = rows[0]
        non_empty_original = [
            v for k, v in first.items()
            if k in original_field_set
            and v is not None
            and str(v).strip() != ""
        ]
        if non_empty_original and all(_is_numeric_like(v) for v in non_empty_original):
            rows = rows[1:]

    # Build a minimal display_columns list (original fields only — export
    # schema overrides column order anyway, but we need it for apply_derived_columns).
    display_columns = list(sheet_dataset.field_names)

    meta_mosad_id = sheet_dataset.get_metadata("MosadID")
    rows, display_columns = apply_derived_columns(
        rows=rows,
        field_names=sheet_dataset.field_names,
        display_columns=display_columns,
        meta_mosad_id=meta_mosad_id,
    )

    return rows, display_columns


class ExportService:
    """Writes the current in-memory WorkbookDataset to an Excel file for download."""

    def __init__(self, session_service: SessionService, output_dir: Path) -> None:
        self.session_service = session_service
        self.output_dir = output_dir

    def export(self, session_id: str) -> Path:
        """Export the session's workbook using the fixed 14-column schema.

        One worksheet is created per source sheet.  Each worksheet:
        - Has the fixed 14-column header row (EXPORT_HEADERS).
        - Is set to right-to-left sheet direction.
        - Contains only corrected field values (no fallback to originals).
        - Leaves cells blank when a corrected field is absent or empty.

        Returns:
            Path to the exported .xlsx file.

        Raises:
            HTTPException 404: session not found
            HTTPException 500: export failed (session state preserved)
        """
        record = self.session_service.get(session_id)

        # Auto-load all sheets from disk if not yet extracted
        if record.workbook_dataset is None:
            try:
                from src.excel_normalization.io_layer.excel_to_json_extractor import ExcelToJsonExtractor
                from src.excel_normalization.io_layer.excel_reader import ExcelReader
                extractor = ExcelToJsonExtractor(
                    ExcelReader(), skip_empty_rows=False,
                    handle_formulas=True, preserve_types=True,
                )
                wbd = extractor.extract_workbook_to_json(record.working_copy_path)
                self.session_service.update(session_id, workbook_dataset=wbd)
                record = self.session_service.get(session_id)
            except Exception as exc:
                logger.error(f"Failed to load workbook for export: {exc}", exc_info=True)
                raise HTTPException(
                    status_code=500,
                    detail="No workbook data available to export. Please upload a file first.",
                )

        output_filename = _build_export_filename(record)

        self.output_dir.mkdir(parents=True, exist_ok=True)

        # F-06: Delete previous exports for this session's source file stem so
        # the output directory does not grow unboundedly.  Only files matching
        # the exact stem pattern are removed; other files are left untouched.
        original_stem = Path(record.original_filename).stem
        for old_file in self.output_dir.glob(f"{original_stem}_normalized_*.xlsx"):
            try:
                old_file.unlink()
                logger.debug(f"Removed previous export: {old_file.name}")
            except Exception as exc:
                logger.warning(f"Could not remove old export file {old_file}: {exc}")

        output_path = self.output_dir / output_filename

        try:
            wb = Workbook()
            # Remove the default empty sheet openpyxl creates
            if wb.sheetnames:
                wb.remove(wb[wb.sheetnames[0]])

            for sheet_dataset in record.workbook_dataset.sheets:
                export_name = canonical_sheet_name(sheet_dataset.sheet_name)
                ws = wb.create_sheet(title=export_name)

                # Right-to-left sheet direction
                ws.sheet_view.rightToLeft = True

                # Determine schema for this sheet type
                schema = headers_for_sheet(export_name)

                # Header row
                for col_idx, header in enumerate(schema, start=1):
                    cell = ws.cell(row=1, column=col_idx, value=header)
                    cell.alignment = Alignment(horizontal="right")

                # Data rows — same visible rows and derived columns as the UI.
                # visible_rows() applies all filters + serial/MosadID injection.
                data_rows, _ui_cols = visible_rows(sheet_dataset)

                # Inject session-level MosadID and SugMosad into every row
                # (overriding any per-sheet metadata or row-level values).
                # SugMosad uses the first user-entered mosad_type (active default).
                active_mosad_type = record.mosad_types[0] if record.mosad_types else ""
                for row in data_rows:
                    if record.mosad_id:
                        row["MosadID"] = record.mosad_id
                    if active_mosad_type:
                        row["SugMosad"] = active_mosad_type

                # Resolve the serial-number source key for this sheet.
                serial_field = detect_serial_field(sheet_dataset.field_names) or SYNTHETIC_SERIAL_KEY

                out_row = 2
                for row in data_rows:
                    for col_idx, header in enumerate(schema, start=1):
                        json_key = EXPORT_MAPPING.get(header)
                        if json_key is None:
                            continue
                        v = row.get(json_key)
                        if v is not None and v != "":
                            ws.cell(row=out_row, column=col_idx, value=v)
                    out_row += 1

            wb.save(str(output_path))
            logger.info(f"Export successful: {output_path}")

        except Exception as exc:
            logger.error(f"Export failed for session {session_id}: {exc}", exc_info=True)
            raise HTTPException(
                status_code=500,
                detail="Export failed. Please try again. Your session data is preserved.",
            )

        return output_path
