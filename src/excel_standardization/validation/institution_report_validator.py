"""Workbook validation for institution-report fields.

The validator runs after standardization, reads corrected values when present,
and records row-level findings for required fields, duplicate IDs, and date
constraints. Checks that depend on external reference data are documented but
not implemented because the project does not include those sources.
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Dict, List, Optional, Set

from src.excel_standardization.normalized_row_contract import (
    select_first_present,
    validation_source_value,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Sheet name constants
# ---------------------------------------------------------------------------
SHEET_ANASHEY_TZEVET = "AnasheyTzevet"
SHEET_DAYARIM_YAHIDIM = "DayarimYahidim"
SHEET_MESHKEY_BAYT = "MeshkeyBayt"

KNOWN_SHEETS = {SHEET_ANASHEY_TZEVET, SHEET_DAYARIM_YAHIDIM, SHEET_MESHKEY_BAYT}

# ---------------------------------------------------------------------------
# Status message constants
# ---------------------------------------------------------------------------

# MosadID
MSG_MOSAD_ID_MISSING = "חסר מספר מוסד"

# SugMosad
MSG_SUG_MOSAD_MISSING = "חסר סוג מוסד"
MSG_SUG_MOSAD_NOT_NUMERIC = "סוג מוסד חייב להכיל ספרות בלבד"
MSG_SUG_MOSAD_TOO_SHORT = "סוג מוסד חייב להכיל לפחות 3 ספרות"

# MisparDiraBeMosad
MSG_DIRA_NOT_NUMERIC = "מספר דירה במוסד חייב להכיל ספרות בלבד"

# ShemPrati
MSG_SHEM_PRATI_MISSING = "חסר שם פרטי"

# ShemMishpaha
MSG_SHEM_MISHPAHA_MISSING = "חסר שם משפחה"

# MisparZehut
MSG_MISPAR_ZEHUT_MISSING = "חסר מספר זהות"
MSG_MISPAR_ZEHUT_DUPLICATE_SHEET = "מספר זהות כפול בגיליון הנוכחי"
MSG_MISPAR_ZEHUT_DUPLICATE_WORKBOOK = "מספר זהות כפול בחוברת העבודה"

# Min (gender)
MSG_MIN_INVALID = "קוד מין לא תקין - חייב להיות 1 (זכר) או 2 (נקבה)"

# Birth date
MSG_SHNAT_LIDA_MISSING = "חסרה שנת לידה"
MSG_SHNAT_LIDA_NOT_NUMERIC = "שנת לידה חייבת להכיל ספרות בלבד"
MSG_SHNAT_LIDA_TOO_EARLY = "שנת לידה לפני 1906"
MSG_SHNAT_LIDA_FUTURE = "שנת לידה עתידית"
MSG_HODESH_LIDA_MISSING = "חסר חודש לידה"
MSG_HODESH_LIDA_NOT_NUMERIC = "חודש לידה חייב להכיל ספרות בלבד"
MSG_HODESH_LIDA_RANGE = "חודש לידה חייב להיות בין 1 ל-12"
MSG_YOM_LIDA_MISSING = "חסר יום לידה"
MSG_YOM_LIDA_NOT_NUMERIC = "יום לידה חייב להכיל ספרות בלבד"
MSG_YOM_LIDA_RANGE = "יום לידה לא תקין לחודש ולשנה שנבחרו"

# Entry date
MSG_SHNAT_KNISA_MISSING = "חסרה שנת כניסה"
MSG_SHNAT_KNISA_NOT_NUMERIC = "שנת כניסה חייבת להכיל ספרות בלבד"
MSG_SHNAT_KNISA_AFTER_CENSUS = "תאריך כניסה מאוחר מהתאריך שנקבע לדיווח"
MSG_SHNAT_KNISA_BEFORE_BIRTH = "תאריך כניסה לפני תאריך לידה"
MSG_HODESH_KNISA_MISSING = "חסר חודש כניסה"
MSG_HODESH_KNISA_NOT_NUMERIC = "חודש כניסה חייב להכיל ספרות בלבד"
MSG_HODESH_KNISA_RANGE = "חודש כניסה חייב להיות בין 1 ל-12"
MSG_YOM_KNISA_MISSING = "חסר יום כניסה"
MSG_YOM_KNISA_NOT_NUMERIC = "יום כניסה חייב להכיל ספרות בלבד"
MSG_YOM_KNISA_RANGE = "יום כניסה לא תקין לחודש ולשנה שנבחרו"


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

# המחלקה מייצגת ממצא validation יחיד בשדה אחד בשורת דיווח.
@dataclass
class ValidationResult:
    """A single validation finding for one field in one row."""
    field_name: str
    message: str
    severity: str = "error"   # "error" | "warning"

    # הפונקציה מציגה ממצא validation בפורמט קריא ללוגים ודוחות.
    def __str__(self) -> str:
        return f"[{self.severity.upper()}] {self.field_name}: {self.message}"


# המחלקה אוספת את כל ממצאי ה־validation עבור שורה אחת.
@dataclass
class RowValidationResult:
    """Aggregated validation results for a single data row."""
    row_index: int                              # 0-based index within the sheet rows list
    row_uid: Optional[str]                      # _row_uid if present
    findings: List[ValidationResult] = field(default_factory=list)

    @property
    # הפונקציה מחזירה האם לשורה אין שגיאות חסומות.
    def is_valid(self) -> bool:
        return not any(f.severity == "error" for f in self.findings)

    @property
    # הפונקציה מחזירה האם לשורה יש אזהרות שאינן חוסמות.
    def has_warnings(self) -> bool:
        return any(f.severity == "warning" for f in self.findings)

    # הפונקציה מוסיפה ממצא validation לשורה במהלך בדיקת החוקים.
    def add(self, field_name: str, message: str, severity: str = "error") -> None:
        self.findings.append(ValidationResult(field_name=field_name, message=message, severity=severity))

    # הפונקציה מאחדת הודעות validation לטקסט קצר שנשמר על השורה.
    def status_summary(self) -> str:
        """Return a pipe-separated summary of all messages, or empty string."""
        return " | ".join(f.message for f in self.findings)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# הפונקציה ממירה ערך לטקסט אחיד לפני בדיקות חובה וטווחים.
def _to_str(value: Any) -> str:
    """Safely convert any value to a stripped string."""
    if value is None:
        return ""
    return str(value).strip()


# הפונקציה בודקת אם ערך טקסטואלי מורכב מספרות בלבד.
def _is_numeric_str(value: str) -> bool:
    """Return True if the stripped string is all digits (non-empty)."""
    return bool(value) and value.isdigit()


# הפונקציה ממירה ערך למספר שלם בלי להפיל את תהליך ה־validation.
def _to_int_safe(value: Any) -> Optional[int]:
    """Convert value to int, return None on failure."""
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        return None


# הפונקציה קוראת את הערך הראשון שקיים מבין כמה שמות שדה אפשריים.
def _get_field(row: Dict[str, Any], *keys: str) -> Any:
    """Return the first non-None, non-empty value found among the given keys."""
    return select_first_present(row, keys)


# הפונקציה מעדיפה ערך corrected אך נופלת לערך מקור כשאין תיקון.
def _get_corrected_or_original(row: Dict[str, Any], base_name: str) -> Any:
    """Return corrected value if present, else original."""
    return validation_source_value(row, base_name)


# ---------------------------------------------------------------------------
# Main validator
# ---------------------------------------------------------------------------

# המחלקה מייצגת שכבת validation עסקית שרצה אחרי סטנדרטיזציה על גיליונות מוסד מוכרים.
class InstitutionReportValidator:
    """Validates institution-report rows against mandatory field requirements.

    Usage (single sheet):
        validator = InstitutionReportValidator(sheet_name="DayarimYahidim")
        results = validator.validate_sheet(rows)

    Usage (full workbook — enables cross-sheet duplicate detection):
        validator = InstitutionReportValidator()
        workbook_results = validator.validate_workbook(sheets_dict)
        # sheets_dict: {sheet_name: [row_dict, ...]}

    The validator writes validation status back into each row dict under the
    key ``_validation_status`` (pipe-separated messages) and ``_validation_ok``
    (bool).  This allows downstream export/display to surface the messages.
    """

    # Minimum birth year per requirements
    MIN_BIRTH_YEAR = 1906

    # הפונקציה מאתחלת הקשר validation, כולל שם גיליון ושנת דיווח.
    def __init__(
        self,
        sheet_name: Optional[str] = None,
        census_year: Optional[int] = None,
    ) -> None:
        """
        Args:
            sheet_name: If set, sheet-specific rules (e.g. YomKnisa required
                        only for DayarimYahidim) are applied.  Pass None when
                        the sheet name is unknown or when calling validate_workbook.
            census_year: The reporting/census year used as the entry-date cutoff.
                         Defaults to current_year - 1 (matching DateEngine behavior).
                         Override via CENSUS_YEAR environment variable or this param.
        """
        self.sheet_name = sheet_name

        if census_year is not None:
            self.census_year = census_year
        else:
            env_year = os.environ.get("CENSUS_YEAR")
            if env_year and env_year.isdigit():
                self.census_year = int(env_year)
            else:
                self.census_year = date.today().year - 1

        logger.debug(
            "InstitutionReportValidator initialized: sheet=%s, census_year=%d",
            sheet_name,
            self.census_year,
        )

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    # הפונקציה מריצה validation על כל הגיליונות ומאפשרת בדיקות רוחביות כמו כפילויות.
    def validate_workbook(
        self,
        sheets: Dict[str, List[Dict[str, Any]]],
        sheet_metadata: Optional[Dict[str, Dict[str, Any]]] = None,
    ) -> Dict[str, List[RowValidationResult]]:
        """Validate all sheets in a workbook, including cross-sheet duplicate detection.

        Args:
            sheets: Mapping of sheet_name → list of row dicts.
            sheet_metadata: Optional mapping of sheet_name → metadata dict.
                            Supports keys: "MosadID", "SugMosad".

        Returns:
            Mapping of sheet_name → list of RowValidationResult.
            Each row dict is mutated in-place with ``_validation_status`` and
            ``_validation_ok`` keys.
        """
        # Collect all corrected MisparZehut values across the workbook for cross-sheet check.
        # Only use non-empty corrected IDs — empty corrected_id means the engine rejected
        # the value (invalid checksum, moved to passport, etc.).
        workbook_id_registry: Dict[str, List[str]] = {}  # id_value → [sheet_name, ...]
        for sname, rows in sheets.items():
            for row in rows:
                id_val = _to_str(row.get("id_number_corrected") or "")
                if id_val:
                    workbook_id_registry.setdefault(id_val, []).append(sname)

        results: Dict[str, List[RowValidationResult]] = {}
        for sname, rows in sheets.items():
            meta = (sheet_metadata or {}).get(sname, {})
            sheet_validator = InstitutionReportValidator(
                sheet_name=sname,
                census_year=self.census_year,
            )
            results[sname] = sheet_validator.validate_sheet(
                rows,
                workbook_id_registry=workbook_id_registry,
                sheet_mosad_id=meta.get("MosadID"),
                sheet_sug_mosad=meta.get("SugMosad"),
            )

        return results

    # הפונקציה מריצה validation על גיליון אחד ומעדכנת סטטוס בכל שורה.
    def validate_sheet(
        self,
        rows: List[Dict[str, Any]],
        workbook_id_registry: Optional[Dict[str, List[str]]] = None,
        sheet_mosad_id: Optional[str] = None,
        sheet_sug_mosad: Optional[str] = None,
    ) -> List[RowValidationResult]:
        """Validate all rows in a single sheet.

        Args:
            rows: List of row dicts (may contain _corrected fields from pipeline).
            workbook_id_registry: Optional mapping of id_value → [sheet_names]
                                  for cross-workbook duplicate detection.
            sheet_mosad_id: MosadID from sheet/session metadata (injected when
                            the row dicts don't carry it yet, e.g. before export).
            sheet_sug_mosad: SugMosad from sheet/session metadata (same reason).

        Returns:
            List of RowValidationResult, one per row.
            Each row dict is mutated in-place with ``_validation_status`` and
            ``_validation_ok`` keys.
        """
        # Build within-sheet ID set for duplicate detection.
        # Only use the corrected ID value (id_number_corrected) when it is
        # non-empty — an empty corrected ID means the IdentifierEngine rejected
        # the value (invalid checksum, moved to passport, etc.) and we must not
        # flag those as duplicates.
        sheet_id_seen: Dict[str, int] = {}  # id_value → first row index (0-based)
        for idx, row in enumerate(rows):
            id_val = _to_str(row.get("id_number_corrected") or "")
            if id_val and id_val not in sheet_id_seen:
                sheet_id_seen[id_val] = idx

        results: List[RowValidationResult] = []
        for idx, row in enumerate(rows):
            row_uid = row.get("_row_uid")
            result = RowValidationResult(row_index=idx, row_uid=row_uid)

            # Inject sheet-level MosadID/SugMosad into a temporary view of the
            # row so the validator can check them even when they are not yet in
            # the row dict (they are injected by the export/UI layer later).
            effective_row = row
            if sheet_mosad_id is not None and not _get_field(row, "MosadID", "mosad_id"):
                effective_row = dict(row)
                effective_row["MosadID"] = sheet_mosad_id
            if sheet_sug_mosad is not None and not _get_field(effective_row, "SugMosad", "sug_mosad"):
                if effective_row is row:
                    effective_row = dict(row)
                effective_row["SugMosad"] = sheet_sug_mosad

            self._validate_mosad_id(effective_row, result)
            self._validate_sug_mosad(effective_row, result)
            self._validate_mispar_dira(effective_row, result)
            self._validate_shem_prati(effective_row, result)
            self._validate_shem_mishpaha(effective_row, result)
            self._validate_mispar_zehut(
                effective_row, result, idx, sheet_id_seen, workbook_id_registry
            )
            self._validate_min(effective_row, result)
            self._validate_birth_date(effective_row, result)
            self._validate_entry_date(effective_row, result)

            # Write status back into the ORIGINAL row dict so downstream can surface it.
            row["_validation_status"] = result.status_summary()
            row["_validation_ok"] = result.is_valid

            results.append(result)

        return results

    # הפונקציה מריצה את כל בדיקות השורה ומחזירה אוסף ממצאים.
    def validate_row(
        self,
        row: Dict[str, Any],
        row_index: int = 0,
        sheet_id_seen: Optional[Dict[str, int]] = None,
        workbook_id_registry: Optional[Dict[str, List[str]]] = None,
        sheet_mosad_id: Optional[str] = None,
        sheet_sug_mosad: Optional[str] = None,
    ) -> RowValidationResult:
        """Validate a single row.  Useful for unit tests and one-off checks."""
        row_uid = row.get("_row_uid")
        result = RowValidationResult(row_index=row_index, row_uid=row_uid)

        # Inject sheet-level MosadID/SugMosad if not in row
        effective_row = row
        if sheet_mosad_id is not None and not _get_field(row, "MosadID", "mosad_id"):
            effective_row = dict(row)
            effective_row["MosadID"] = sheet_mosad_id
        if sheet_sug_mosad is not None and not _get_field(effective_row, "SugMosad", "sug_mosad"):
            if effective_row is row:
                effective_row = dict(row)
            effective_row["SugMosad"] = sheet_sug_mosad

        self._validate_mosad_id(effective_row, result)
        self._validate_sug_mosad(effective_row, result)
        self._validate_mispar_dira(effective_row, result)
        self._validate_shem_prati(effective_row, result)
        self._validate_shem_mishpaha(effective_row, result)
        self._validate_mispar_zehut(
            effective_row, result, row_index,
            sheet_id_seen or {},
            workbook_id_registry,
        )
        self._validate_min(effective_row, result)
        self._validate_birth_date(effective_row, result)
        self._validate_entry_date(effective_row, result)

        row["_validation_status"] = result.status_summary()
        row["_validation_ok"] = result.is_valid

        return result

    # ------------------------------------------------------------------
    # Field validators
    # ------------------------------------------------------------------

    # הפונקציה בודקת שמספר מוסד קיים בשורה או במטא־דאטה של הגיליון.
    def _validate_mosad_id(self, row: Dict[str, Any], result: RowValidationResult) -> None:
        """MosadID: required in export only; missing values are reported."""
        val = _to_str(_get_field(row, "MosadID", "mosad_id"))
        if not val:
            result.add("MosadID", MSG_MOSAD_ID_MISSING)

    # הפונקציה בודקת שסוג מוסד קיים ותקין לפני יצוא.
    def _validate_sug_mosad(self, row: Dict[str, Any], result: RowValidationResult) -> None:
        """SugMosad is required, numeric, and at least three digits long.

        The source project does not include a SugMosad dictionary, so
        membership checks are intentionally not implemented.
        """
        val = _to_str(_get_field(row, "SugMosad", "sug_mosad"))
        if not val:
            result.add("SugMosad", MSG_SUG_MOSAD_MISSING)
            return
        if not _is_numeric_str(val):
            result.add("SugMosad", MSG_SUG_MOSAD_NOT_NUMERIC)
            return
        if len(val) < 3:
            result.add("SugMosad", MSG_SUG_MOSAD_TOO_SHORT)

    # הפונקציה בודקת מספר דירה בגיליונות שבהם השדה נדרש.
    def _validate_mispar_dira(self, row: Dict[str, Any], result: RowValidationResult) -> None:
        """MisparDiraBeMosad: optional; if provided must be numeric.

        Relevant for AnasheyTzevet and MeshkeyBayt only.
        DayarimYahidim does not have this field — skip silently.
        """
        if self.sheet_name == SHEET_DAYARIM_YAHIDIM:
            return
        val = _to_str(_get_field(row, "MisparDiraBeMosad"))
        if not val:
            return  # optional — empty is fine
        if not _is_numeric_str(val):
            result.add("MisparDiraBeMosad", MSG_DIRA_NOT_NUMERIC)

    # הפונקציה בודקת ששדה שם פרטי קיים לאחר נרמול.
    def _validate_shem_prati(self, row: Dict[str, Any], result: RowValidationResult) -> None:
        """ShemPrati: required, must not be empty after normalization."""
        # Prefer corrected value; fall back to original.
        val = _to_str(_get_corrected_or_original(row, "first_name"))
        if not val:
            result.add("ShemPrati", MSG_SHEM_PRATI_MISSING)

    # הפונקציה בודקת ששדה שם משפחה קיים לאחר נרמול.
    def _validate_shem_mishpaha(self, row: Dict[str, Any], result: RowValidationResult) -> None:
        """ShemMishpaha: required, must not be empty after normalization."""
        val = _to_str(_get_corrected_or_original(row, "last_name"))
        if not val:
            result.add("ShemMishpaha", MSG_SHEM_MISHPAHA_MISSING)

    # הפונקציה בודקת מספר זהות, כולל כפילויות בגיליון וב־workbook.
    def _validate_mispar_zehut(
        self,
        row: Dict[str, Any],
        result: RowValidationResult,
        row_index: int,
        sheet_id_seen: Dict[str, int],
        workbook_id_registry: Optional[Dict[str, List[str]]],
    ) -> None:
        """MisparZehut is required, must be valid, and must be unique.

        Missing values are errors. Duplicate values are checked within each
        sheet and across the workbook. Registry and related-institution checks
        are intentionally not implemented because the required data is absent.
        """
        id_val = _to_str(row.get("id_number_corrected") or "")

        if not id_val:
            # If the corrected ID is empty, keep only the missing-field check.
            original_id = _to_str(row.get("id_number") or "")
            if not original_id:
                result.add("MisparZehut", MSG_MISPAR_ZEHUT_MISSING)
            # A rejected ID already has a status from IdentifierEngine.
            return

        # Duplicate within the current sheet.
        first_occurrence = sheet_id_seen.get(id_val)
        if first_occurrence is not None and first_occurrence != row_index:
            result.add("MisparZehut", MSG_MISPAR_ZEHUT_DUPLICATE_SHEET)

        # Duplicate across the workbook.
        if workbook_id_registry is not None:
            sheets_with_id = workbook_id_registry.get(id_val, [])
            if len(sheets_with_id) > 1:
                result.add("MisparZehut", MSG_MISPAR_ZEHUT_DUPLICATE_WORKBOOK, severity="warning")


    # הפונקציה בודקת שקוד המגדר המתוקן הוא אחד מהערכים המותרים.
    def _validate_min(self, row: Dict[str, Any], result: RowValidationResult) -> None:
        """Min (gender): if present, must be 1 or 2.

        The GenderEngine normalizes textual values to 1/2 or "".
        An empty string from the engine means the value was unrecognized.
        We flag that as an error here.
        """
        # Check corrected value first (set by GenderEngine via pipeline).
        corrected = row.get("gender_corrected")
        if corrected is None:
            # No corrected field — check original.
            original = row.get("gender")
            if original is None or str(original).strip() == "":
                return  # empty is allowed (field is optional per requirements)
            # Original present but no corrected — treat as unvalidated, skip.
            return

        corrected_str = str(corrected).strip()
        if corrected_str == "":
            # GenderEngine returned "" — unrecognized value.
            result.add("Min", MSG_MIN_INVALID)
            return

        try:
            code = int(float(corrected_str))
            if code not in (1, 2):
                result.add("Min", MSG_MIN_INVALID)
        except (TypeError, ValueError):
            result.add("Min", MSG_MIN_INVALID)

    # הפונקציה בודקת רכיבי תאריך לידה, חובה וטווחים עסקיים.
    def _validate_birth_date(self, row: Dict[str, Any], result: RowValidationResult) -> None:
        """Birth date is checked for required values, numeric input, range, and
        the 1906 minimum year.

        The DateEngine also enforces these rules and writes birth_date_status.
        This validator keeps a second pass for raw or uncorrected values.
        """
        year_val = _get_corrected_or_original(row, "birth_year")
        month_val = _get_corrected_or_original(row, "birth_month")
        day_val = _get_corrected_or_original(row, "birth_day")

        year_str = _to_str(year_val)
        month_str = _to_str(month_val)
        day_str = _to_str(day_val)

        # Required checks
        if not year_str:
            result.add("ShnatLida", MSG_SHNAT_LIDA_MISSING)
        elif not _is_numeric_str(year_str):
            result.add("ShnatLida", MSG_SHNAT_LIDA_NOT_NUMERIC)
        else:
            yr = _to_int_safe(year_str)
            if yr is not None:
                today = date.today()
                if yr < self.MIN_BIRTH_YEAR:
                    result.add("ShnatLida", MSG_SHNAT_LIDA_TOO_EARLY)
                elif yr > today.year:
                    result.add("ShnatLida", MSG_SHNAT_LIDA_FUTURE)

        if not month_str:
            result.add("HodeshLida", MSG_HODESH_LIDA_MISSING)
        elif not _is_numeric_str(month_str):
            result.add("HodeshLida", MSG_HODESH_LIDA_NOT_NUMERIC)
        else:
            mo = _to_int_safe(month_str)
            if mo is not None and not (1 <= mo <= 12):
                result.add("HodeshLida", MSG_HODESH_LIDA_RANGE)

        if not day_str:
            result.add("YomLida", MSG_YOM_LIDA_MISSING)
        elif not _is_numeric_str(day_str):
            result.add("YomLida", MSG_YOM_LIDA_NOT_NUMERIC)
        else:
            dy = _to_int_safe(day_str)
            if dy is not None and not (1 <= dy <= 31):
                result.add("YomLida", MSG_YOM_LIDA_RANGE)

    # הפונקציה בודקת רכיבי תאריך כניסה מול שנת הדיווח ותאריך הלידה.
    def _validate_entry_date(self, row: Dict[str, Any], result: RowValidationResult) -> None:
        """Entry date is checked for numeric values, range, and the census cutoff.

        YomKnisa is required only for DayarimYahidim. The minimum-entry-age
        rule is not implemented because the required SugMosad reference data is
        not available in this project.
        """
        year_val = _get_corrected_or_original(row, "entry_year")
        month_val = _get_corrected_or_original(row, "entry_month")
        day_val = _get_corrected_or_original(row, "entry_day")

        year_str = _to_str(year_val)
        month_str = _to_str(month_val)
        day_str = _to_str(day_val)

        # Required checks
        if not year_str:
            result.add("shnatknisa", MSG_SHNAT_KNISA_MISSING)
        elif not _is_numeric_str(year_str):
            result.add("shnatknisa", MSG_SHNAT_KNISA_NOT_NUMERIC)
        else:
            yr = _to_int_safe(year_str)
            if yr is not None and yr > self.census_year:
                result.add("shnatknisa", MSG_SHNAT_KNISA_AFTER_CENSUS)

        if not month_str:
            result.add("Hodeshknisa", MSG_HODESH_KNISA_MISSING)
        elif not _is_numeric_str(month_str):
            result.add("Hodeshknisa", MSG_HODESH_KNISA_NOT_NUMERIC)
        else:
            mo = _to_int_safe(month_str)
            if mo is not None and not (1 <= mo <= 12):
                result.add("Hodeshknisa", MSG_HODESH_KNISA_RANGE)

        # YomKnisa is required only for DayarimYahidim.
        yom_required = (self.sheet_name == SHEET_DAYARIM_YAHIDIM)
        if not day_str:
            if yom_required:
                result.add("YomKnisa", MSG_YOM_KNISA_MISSING)
            # else: optional for other sheets — skip
        elif not _is_numeric_str(day_str):
            result.add("YomKnisa", MSG_YOM_KNISA_NOT_NUMERIC)
        else:
            dy = _to_int_safe(day_str)
            if dy is not None and not (1 <= dy <= 31):
                result.add("YomKnisa", MSG_YOM_KNISA_RANGE)

        # When available, compute age = entry_year - birth_year and compare
        # against the dictionary value for the row's SugMosad.
