"""NormalizationPipeline: Apply normalization engines to JSON rows.

This module provides the NormalizationPipeline class that orchestrates the
application of normalization engines (NameEngine, GenderEngine, DateEngine,
IdentifierEngine) to JSON row data extracted from Excel worksheets.

The pipeline operates on JSON data structures, maintaining clean separation
between IO operations and business logic. It preserves original values while
creating corrected fields with the "_corrected" suffix.
"""

import logging
from typing import Optional, Dict, Any, List, Tuple
from ..data_types import JsonRow, SheetDataset
from ..engines.name_engine import NameEngine
from ..engines.gender_engine import GenderEngine
from ..engines.date_engine import DateEngine
from ..engines.identifier_engine import IdentifierEngine
from ..engines.text_processor import TextProcessor

# Configure logger for this module
logger = logging.getLogger(__name__)


def _detect_date_format_pattern(rows: List[JsonRow]) -> "DateFormatPattern":
    """F-03: Detect whether date values in this dataset use DDMM or MMDD ordering.

    Samples the first 20 rows looking for separated date strings (containing
    "/" or ".") in any date-related field.  Counts how many values have a
    first part > 12 (unambiguously DD) vs a second part > 12 (unambiguously MM).
    Falls back to DDMM when the evidence is inconclusive.

    Args:
        rows: List of JSON row dicts from the dataset.

    Returns:
        DateFormatPattern.MMDD if MMDD evidence outweighs DDMM evidence,
        DateFormatPattern.DDMM otherwise.
    """
    from ..data_types import DateFormatPattern

    date_fields = (
        "birth_date", "entry_date",
        "birth_year", "entry_year",  # sometimes a full date string lands here
    )
    ddmm = 0
    mmdd = 0

    for row in rows[:20]:
        for field in date_fields:
            val = row.get(field)
            if not val or not isinstance(val, str):
                continue
            s = val.replace(".", "/")
            if "/" not in s:
                continue
            parts = s.split("/")
            if len(parts) < 2:
                continue
            try:
                a, b = int(parts[0]), int(parts[1])
                if a > 12 and b <= 12:
                    ddmm += 1
                elif b > 12 and a <= 12:
                    mmdd += 1
            except (ValueError, TypeError):
                pass

    return DateFormatPattern.MMDD if mmdd > ddmm else DateFormatPattern.DDMM


class NormalizationPipeline:
    """Apply normalization engines to JSON rows.
    
    This class orchestrates the application of normalization engines to JSON
    row data. It acts as an adapter between JSON data structures and the
    existing normalization engines, which operate on string and numeric values.
    
    The pipeline:
    - Accepts JSON rows with original field values
    - Applies configured normalization engines
    - Creates corrected fields with "_corrected" suffix
    - Preserves original values (non-destructive)
    - Handles missing fields and engine failures gracefully
    
    Attributes:
        name_engine: Engine for normalizing name fields
        gender_engine: Engine for normalizing gender values
        date_engine: Engine for parsing and validating dates
        identifier_engine: Engine for validating ID and passport values
        apply_name_normalization_enabled: Whether to apply name normalization
        apply_gender_normalization_enabled: Whether to apply gender normalization
        apply_date_normalization_enabled: Whether to apply date normalization
        apply_identifier_normalization_enabled: Whether to apply identifier normalization
    
    Example:
        # Create pipeline with all engines
        pipeline = NormalizationPipeline(
            name_engine=NameEngine(TextProcessor()),
            gender_engine=GenderEngine(),
            date_engine=DateEngine(),
            identifier_engine=IdentifierEngine()
        )
        
        # Normalize a single row
        row = {"first_name": "יוסי", "gender": "ז", "id_number": "123456789"}
        normalized_row = pipeline.normalize_row(row)
        # Result: {
        #     "first_name": "יוסי",
        #     "first_name_corrected": "יוסי",
        #     "gender": "ז",
        #     "gender_corrected": 2,
        #     "id_number": "123456789",
        #     "id_number_corrected": "123456789"
        # }
        
        # Normalize an entire dataset
        dataset = SheetDataset(...)
        corrected_dataset = pipeline.normalize_dataset(dataset)
    
    Requirements:
        - Validates: Requirements 12.1, 17.6
    """
    
    def __init__(
        self,
        name_engine: Optional[NameEngine] = None,
        gender_engine: Optional[GenderEngine] = None,
        date_engine: Optional[DateEngine] = None,
        identifier_engine: Optional[IdentifierEngine] = None,
        apply_name_normalization_enabled: bool = True,
        apply_gender_normalization_enabled: bool = True,
        apply_date_normalization_enabled: bool = True,
        apply_identifier_normalization_enabled: bool = True
    ):
        """Initialize NormalizationPipeline with engine dependencies.
        
        Args:
            name_engine: Engine for normalizing name fields (optional)
            gender_engine: Engine for normalizing gender values (optional)
            date_engine: Engine for parsing and validating dates (optional)
            identifier_engine: Engine for validating ID and passport values (optional)
            apply_name_normalization_enabled: Whether to apply name normalization
            apply_gender_normalization_enabled: Whether to apply gender normalization
            apply_date_normalization_enabled: Whether to apply date normalization
            apply_identifier_normalization_enabled: Whether to apply identifier normalization
        
        Note:
            If an engine is not provided, the corresponding normalization will be
            skipped even if the enabled flag is True. This allows for flexible
            configuration where only specific engines are used.
        
        Requirements:
            - Validates: Requirements 12.1, 17.6
        """
        self.name_engine = name_engine
        self.gender_engine = gender_engine
        self.date_engine = date_engine
        self.identifier_engine = identifier_engine
        
        # Configuration flags for which engines to apply
        self.apply_name_normalization_enabled = apply_name_normalization_enabled
        self.apply_gender_normalization_enabled = apply_gender_normalization_enabled
        self.apply_date_normalization_enabled = apply_date_normalization_enabled
        self.apply_identifier_normalization_enabled = apply_identifier_normalization_enabled
    
    def normalize_row(self, json_row: JsonRow, row_number: Optional[int] = None) -> JsonRow:
        """Apply normalization engines to a single row.
        
        Creates corrected fields for each normalized value. Original values
        are never modified. Corrected fields use the "_corrected" suffix.
        
        Args:
            json_row: Dictionary with original field values
            row_number: Optional row number for error logging (1-based)
        
        Returns:
            Dictionary with original and corrected field values
        
        Example:
            row = {"first_name": "יוסי", "gender": "ז"}
            normalized = pipeline.normalize_row(row, row_number=5)
            # Result: {
            #     "first_name": "יוסי",
            #     "first_name_corrected": "יוסי",
            #     "gender": "ז",
            #     "gender_corrected": 2
            # }
        
        Requirements:
            - Validates: Requirements 12.2, 13.2-13.5, 18.1-18.4
        """
        # Create a copy to avoid modifying the original
        result = json_row.copy()
        
        # Track failed normalizations for this row
        failed_fields: List[str] = []
        
        # Apply each normalization engine
        if self.apply_name_normalization_enabled and self.name_engine:
            failures = self.apply_name_normalization(result, row_number)
            failed_fields.extend(failures)
        
        if self.apply_gender_normalization_enabled and self.gender_engine:
            failures = self.apply_gender_normalization(result, row_number)
            failed_fields.extend(failures)
        
        if self.apply_date_normalization_enabled and self.date_engine:
            failures = self.apply_date_normalization(result, row_number)
            failed_fields.extend(failures)
        
        if self.apply_identifier_normalization_enabled and self.identifier_engine:
            failures = self.apply_identifier_normalization(result, row_number)
            failed_fields.extend(failures)
        
        # Store failed fields in metadata if any failures occurred
        if failed_fields:
            result["_normalization_failures"] = failed_fields
        
        return result
    
    def apply_name_normalization(self, json_row: JsonRow, row_number: Optional[int] = None) -> List[str]:
        """Apply NameEngine to name fields in the row.

        Updates json_row with corrected fields for:
        - first_name  -> first_name_corrected  (with last-name removal if applicable)
        - last_name   -> last_name_corrected
        - father_name -> father_name_corrected (with last-name removal if applicable)

        Last-name removal uses the two-stage logic in NameEngine:
          Stage A: substring removal.
          Stage B: positional fallback — only when Stage A made no change.

        The pattern for father_name and first_name is stored on the pipeline
        instance (set by normalize_dataset before iterating rows).

        Args:
            json_row: Dictionary to update with corrected name fields
            row_number: Optional row number for error logging (1-based)

        Returns:
            List of field names that failed normalization

        Requirements:
            - Validates: Requirements 12.3, 12.8, 14.1-14.5, 18.1-18.4
        """
        from ..data_types import FatherNamePattern

        failed_fields: List[str] = []

        try:
            # --- last_name: clean only, no removal ---
            if "last_name" in json_row:
                original = json_row["last_name"]
                if original is None or original == "":
                    json_row["last_name_corrected"] = original
                else:
                    json_row["last_name_corrected"] = self.name_engine.normalize_name(str(original))

            # Resolve cleaned last name for use in removal below
            cleaned_last = ""
            if "last_name" in json_row:
                raw_last = json_row.get("last_name")
                if raw_last:
                    cleaned_last = self.name_engine.normalize_name(str(raw_last))

            # --- first_name: clean + last-name removal ---
            if "first_name" in json_row:
                original = json_row["first_name"]
                if original is None or original == "":
                    json_row["first_name_corrected"] = original
                else:
                    cleaned = self.name_engine.normalize_name(str(original))
                    if cleaned_last:
                        pattern = getattr(self, "_first_name_pattern", FatherNamePattern.NONE)
                        cleaned = self.name_engine.remove_last_name_from_first_name(
                            cleaned, cleaned_last, pattern
                        )
                    json_row["first_name_corrected"] = cleaned

            # --- father_name: clean + last-name removal ---
            if "father_name" in json_row:
                original = json_row["father_name"]
                if original is None or original == "":
                    json_row["father_name_corrected"] = original
                else:
                    cleaned = self.name_engine.normalize_name(str(original))
                    if cleaned_last:
                        pattern = getattr(self, "_father_name_pattern", FatherNamePattern.NONE)
                        cleaned = self.name_engine.remove_last_name_from_father(
                            cleaned, cleaned_last, pattern
                        )
                    json_row["father_name_corrected"] = cleaned

        except Exception as e:
            for field in ["first_name", "last_name", "father_name"]:
                if field in json_row and f"{field}_corrected" not in json_row:
                    json_row[f"{field}_corrected"] = json_row[field]
                    failed_fields.append(field)
            row_info = f"row {row_number}" if row_number is not None else "unknown row"
            logger.error(f"Name normalization failed at {row_info}: {e}")

        return failed_fields
    
    def apply_gender_normalization(self, json_row: JsonRow, row_number: Optional[int] = None) -> List[str]:
        """Apply GenderEngine to gender field in the row.
        
        Updates json_row with corrected field:
        - gender -> gender_corrected
        
        Args:
            json_row: Dictionary to update with corrected gender field
            row_number: Optional row number for error logging (1-based)
        
        Returns:
            List of field names that failed normalization
        
        Requirements:
            - Validates: Requirements 12.4, 12.8, 14.1-14.5, 18.1-18.4
        """
        failed_fields: List[str] = []
        
        if "gender" in json_row:
            original = json_row["gender"]
            
            # F-04: Handle None, empty string, AND whitespace-only values consistently.
            # Previously only None and "" were caught here; whitespace-only strings
            # fell through to the engine which stripped them and returned 1 (male),
            # inconsistent with how None/"" are handled (preserved as-is).
            if original is None or str(original).strip() == "":
                json_row["gender_corrected"] = original
                return failed_fields
            
            # Try to normalize the gender
            try:
                corrected = self.gender_engine.normalize_gender(original)
                json_row["gender_corrected"] = corrected
            except Exception as e:
                # If engine fails, store original value
                json_row["gender_corrected"] = original
                failed_fields.append("gender")
                
                # Log error with context
                row_info = f"row {row_number}" if row_number is not None else "unknown row"
                logger.error(
                    f"Gender normalization failed for field 'gender' at {row_info}: {str(e)}. "
                    f"Original value: '{original}'"
                )
        
        return failed_fields
    
    def apply_date_normalization(self, json_row: JsonRow, row_number: Optional[int] = None) -> List[str]:
        """Apply DateEngine to date fields in the row.
        
        Updates json_row with corrected fields for:
        - birth_date or birth_year/month/day -> corrected fields
        - entry_date or entry_year/month/day -> corrected fields
        
        Handles both single date fields and split date fields.
        Also cross-validates entry date against birth date (F-02).
        
        Args:
            json_row: Dictionary to update with corrected date fields
            row_number: Optional row number for error logging (1-based)
        
        Returns:
            List of field names that failed normalization
        
        Requirements:
            - Validates: Requirements 12.5, 12.8, 14.1-14.5, 18.1-18.4
        """
        from ..data_types import DateFormatPattern, DateFieldType
        
        failed_fields: List[str] = []
        
        # Process birth date — store result for cross-validation below
        failures, birth_result = self._normalize_date_field(
            json_row,
            "birth",
            DateFieldType.BIRTH_DATE,
            row_number
        )
        failed_fields.extend(failures)
        
        # Process entry date — store result for cross-validation below
        failures, entry_result = self._normalize_date_field(
            json_row,
            "entry",
            DateFieldType.ENTRY_DATE,
            row_number
        )
        failed_fields.extend(failures)

        # F-02: Cross-validate entry date against birth date.
        # DateEngine.validate_entry_before_birth exists but was never called by
        # the pipeline.  We call it here and append a warning to entry_date_status
        # when the entry date precedes the birth date.
        if birth_result is not None and entry_result is not None:
            try:
                if not self.date_engine.validate_entry_before_birth(birth_result, entry_result):
                    warning = "תאריך כניסה לפני תאריך לידה"
                    existing_status = json_row.get("entry_date_status", "")
                    if existing_status:
                        json_row["entry_date_status"] = f"{existing_status} | {warning}"
                    else:
                        json_row["entry_date_status"] = warning
            except Exception:
                pass  # Cross-validation is best-effort; never block normalization
        
        return failed_fields
    
    def _normalize_date_field(self, json_row: JsonRow, prefix: str, field_type, row_number: Optional[int] = None):
        """Helper method to normalize a date field (birth or entry).
        
        Args:
            json_row: Dictionary to update with corrected date fields
            prefix: Field prefix ("birth" or "entry")
            field_type: DateFieldType enum value
            row_number: Optional row number for error logging (1-based)
        
        Returns:
            Tuple of (failed_fields: List[str], date_result: Optional[DateParseResult])
            date_result is the parsed DateParseResult for cross-validation, or None if
            no date fields were present or parsing was skipped.
        """
        from ..data_types import DateFormatPattern
        
        failed_fields: List[str] = []
        date_result = None  # returned for entry-before-birth cross-validation (F-02)
        
        # F-03: Use the per-dataset detected date format pattern instead of always
        # hardcoding DDMM.  The pattern is detected once in normalize_dataset() and
        # cached on the pipeline instance as _date_format_pattern.
        pattern = getattr(self, "_date_format_pattern", DateFormatPattern.DDMM)
        
        # Check for split date fields
        year_field = f"{prefix}_year"
        month_field = f"{prefix}_month"
        day_field = f"{prefix}_day"
        
        # Check for single date field
        date_field = f"{prefix}_date"
        
        # Determine if we have split or single date
        has_split = (year_field in json_row or month_field in json_row or day_field in json_row)
        has_single = date_field in json_row
        
        if has_split:
            # Process split date fields
            year_val = json_row.get(year_field)
            month_val = json_row.get(month_field)
            day_val = json_row.get(day_field)

            # If only year_val is present (month/day are null), treat year_val as main_val.
            # This handles the case where a full date string (ISO, DD/MM/YYYY, etc.) is
            # stored in the year column because the sheet uses a single merged date cell
            # that openpyxl maps to the first split column.
            # Also handle datetime objects stored in the year column.
            from datetime import datetime as _dt, date as _date
            if year_val is not None and month_val is None and day_val is None:
                main_val_for_engine = year_val
                year_val_for_engine = None
                month_val_for_engine = None
                day_val_for_engine = None
            elif isinstance(year_val, (_dt, _date)):
                # Excel datetime stored in year column — treat as main value
                main_val_for_engine = year_val
                year_val_for_engine = None
                month_val_for_engine = None
                day_val_for_engine = None
            else:
                main_val_for_engine = None
                year_val_for_engine = year_val
                month_val_for_engine = month_val
                day_val_for_engine = day_val
            
            try:
                # Use DateEngine to parse from split columns
                result = self.date_engine.parse_date(
                    year_val_for_engine, month_val_for_engine, day_val_for_engine,
                    main_val_for_engine,
                    pattern,
                    field_type
                )
                date_result = result
                
                # Store corrected values
                json_row[f"{year_field}_corrected"] = result.year if result.year is not None else year_val
                json_row[f"{month_field}_corrected"] = result.month if result.month is not None else month_val
                json_row[f"{day_field}_corrected"] = result.day if result.day is not None else day_val
                # Write status text so the UI can display it
                json_row[f"{prefix}_date_status"] = result.status_text
                # Tag whether the year was auto-completed (for list-level majority correction)
                json_row[f"_{prefix}_year_auto_completed"] = result.year_was_auto_completed
                
            except Exception as e:
                # If engine fails, store original values
                json_row[f"{year_field}_corrected"] = year_val
                json_row[f"{month_field}_corrected"] = month_val
                json_row[f"{day_field}_corrected"] = day_val
                json_row[f"{prefix}_date_status"] = ""
                json_row[f"_{prefix}_year_auto_completed"] = False
                
                # Track all three fields as failed
                failed_fields.extend([year_field, month_field, day_field])
                
                # Log error with context
                row_info = f"row {row_number}" if row_number is not None else "unknown row"
                logger.error(
                    f"Date normalization failed for split date fields '{prefix}_*' at {row_info}: {str(e)}. "
                    f"Original values: year={year_val}, month={month_val}, day={day_val}"
                )
        
        elif has_single:
            # Process single date field.
            # The source has one date column (e.g. birth_date) with raw values.
            # We parse the value and ALWAYS write structured year/month/day
            # corrected fields — the same output model as the split path.
            # There is no weaker "leave it as one text field" fallback.
            date_val = json_row.get(date_field)

            # Derive the structured field names (same as split path)
            year_field = f"{prefix}_year"
            month_field = f"{prefix}_month"
            day_field = f"{prefix}_day"

            # Handle None/empty values
            if date_val is None or date_val == "":
                json_row[f"{year_field}_corrected"] = None
                json_row[f"{month_field}_corrected"] = None
                json_row[f"{day_field}_corrected"] = None
                json_row[f"{prefix}_date_status"] = ""
                json_row[f"_{prefix}_year_auto_completed"] = False
                return failed_fields, date_result

            try:
                # Use DateEngine to parse from main value
                result = self.date_engine.parse_date(
                    None, None, None,  # no split values
                    date_val,          # main_val
                    pattern,
                    field_type
                )
                date_result = result

                # Always write structured year/month/day corrected fields.
                # When parsing failed (components are None), write None so the
                # UI shows empty cells rather than the raw unparseable string.
                json_row[f"{year_field}_corrected"] = result.year
                json_row[f"{month_field}_corrected"] = result.month
                json_row[f"{day_field}_corrected"] = result.day
                # Write status text so the UI can display it
                json_row[f"{prefix}_date_status"] = result.status_text
                # Tag whether the year was auto-completed (for list-level majority correction)
                json_row[f"_{prefix}_year_auto_completed"] = result.year_was_auto_completed

            except Exception as e:
                # If engine fails, write empty structured fields
                json_row[f"{year_field}_corrected"] = None
                json_row[f"{month_field}_corrected"] = None
                json_row[f"{day_field}_corrected"] = None
                json_row[f"{prefix}_date_status"] = ""
                json_row[f"_{prefix}_year_auto_completed"] = False
                failed_fields.append(date_field)

                row_info = f"row {row_number}" if row_number is not None else "unknown row"
                logger.error(
                    f"Date normalization failed for field '{date_field}' at {row_info}: {str(e)}. "
                    f"Original value: '{date_val}'"
                )
        
        return failed_fields, date_result
    
    def apply_identifier_normalization(self, json_row: JsonRow, row_number: Optional[int] = None) -> List[str]:
        """Apply IdentifierEngine to identifier fields in the row.
        
        Updates json_row with corrected fields for:
        - id_number -> id_number_corrected
        - passport -> passport_corrected
        
        Args:
            json_row: Dictionary to update with corrected identifier fields
            row_number: Optional row number for error logging (1-based)
        
        Returns:
            List of field names that failed normalization
        
        Requirements:
            - Validates: Requirements 12.6, 12.8, 14.1-14.5, 18.1-18.4
        """
        failed_fields: List[str] = []
        
        # Get original values
        id_value = json_row.get("id_number")
        passport_value = json_row.get("passport")
        
        # Handle case where neither field exists
        if "id_number" not in json_row and "passport" not in json_row:
            return failed_fields
        
        # Handle None/empty values for both fields
        if (id_value is None or id_value == "") and (passport_value is None or passport_value == ""):
            if "id_number" in json_row:
                json_row["id_number_corrected"] = id_value
            if "passport" in json_row:
                json_row["passport_corrected"] = passport_value
            # Always write identifier_status so the column appears consistently
            # in field_names even when both identifiers are empty.
            json_row["identifier_status"] = "חסר מזהים"
            return failed_fields
        
        try:
            # Use IdentifierEngine to normalize both fields together
            result = self.identifier_engine.normalize_identifiers(id_value, passport_value)
            
            # Store corrected values
            if "id_number" in json_row:
                json_row["id_number_corrected"] = result.corrected_id
            if "passport" in json_row:
                json_row["passport_corrected"] = result.corrected_passport
            # Always write the status text so the UI can display it
            json_row["identifier_status"] = result.status_text
                
        except Exception as e:
            # If engine fails, store original values
            if "id_number" in json_row:
                json_row["id_number_corrected"] = id_value
                failed_fields.append("id_number")
            if "passport" in json_row:
                json_row["passport_corrected"] = passport_value
                failed_fields.append("passport")
            json_row["identifier_status"] = ""
            
            # Log error with context
            row_info = f"row {row_number}" if row_number is not None else "unknown row"
            logger.error(
                f"Identifier normalization failed for fields 'id_number'/'passport' at {row_info}: {str(e)}. "
                f"Original values: id_number='{id_value}', passport='{passport_value}'"
            )
        
        return failed_fields
    
    def normalize_dataset(self, raw_dataset: SheetDataset) -> SheetDataset:
        """Apply normalization engines to all rows in dataset.

        Creates a new dataset with both original and corrected values.
        Updates metadata with normalization information and tracks failed normalizations.

        Args:
            raw_dataset: SheetDataset with original values

        Returns:
            SheetDataset with both original and corrected values

        Example:
            raw_dataset = SheetDataset(
                sheet_name="Students",
                header_row=1,
                header_rows_count=1,
                field_names=["first_name", "gender"],
                rows=[
                    {"first_name": "יוסי", "gender": "ז"},
                    {"first_name": "שרה", "gender": "נ"}
                ],
                metadata={}
            )

            corrected_dataset = pipeline.normalize_dataset(raw_dataset)
            # Result: SheetDataset with rows containing both original and corrected fields

        Requirements:
            - Validates: Requirements 12.1-12.2, 13.1-13.7, 18.1-18.4
        """
        # Shallow-copy the dataset shell; rows are already fresh from extraction
        # so a deepcopy is unnecessary and expensive.
        import copy
        from ..data_types import FatherNamePattern
        corrected_dataset = copy.copy(raw_dataset)
        corrected_dataset.rows = list(raw_dataset.rows)   # independent list
        corrected_dataset.metadata = dict(raw_dataset.metadata)

        # ------------------------------------------------------------------
        # Detect last-name removal patterns once per dataset (not per row).
        # Build sample arrays from the first few rows that have both fields.
        # ------------------------------------------------------------------
        if self.apply_name_normalization_enabled and self.name_engine:
            first_sample: List[List] = []
            father_sample: List[List] = []
            last_sample: List[List] = []

            for row in corrected_dataset.rows[:10]:
                fn = row.get("first_name") or ""
                fa = row.get("father_name") or ""
                ln = row.get("last_name") or ""
                if fn and ln:
                    first_sample.append([fn])
                    last_sample.append([ln])
                if fa and ln:
                    father_sample.append([fa])

            # Detect and cache patterns on the pipeline instance so
            # apply_name_normalization can read them per-row.
            self._first_name_pattern = (
                self.name_engine.detect_first_name_pattern(first_sample, last_sample)
                if first_sample else FatherNamePattern.NONE
            )
            self._father_name_pattern = (
                self.name_engine.detect_father_name_pattern(father_sample, last_sample[:len(father_sample)])
                if father_sample else FatherNamePattern.NONE
            )
        else:
            self._first_name_pattern = FatherNamePattern.NONE
            self._father_name_pattern = FatherNamePattern.NONE

        # F-03: Detect date format pattern (DDMM vs MMDD) once per dataset.
        # Previously the pipeline always used DDMM.  Now we sample the first 20
        # rows to detect whether the sheet uses US-style MM/DD dates.
        if self.apply_date_normalization_enabled and self.date_engine:
            self._date_format_pattern = _detect_date_format_pattern(corrected_dataset.rows)
            logger.debug(
                f"Date format pattern detected for sheet '{raw_dataset.sheet_name}': "
                f"{self._date_format_pattern}"
            )
        else:
            from ..data_types import DateFormatPattern
            self._date_format_pattern = DateFormatPattern.DDMM

        # Track normalization statistics
        total_rows = len(corrected_dataset.rows)
        rows_with_failures = 0
        total_field_failures = 0
        failed_rows: List[int] = []

        # Normalize each row
        normalized_rows = []
        for idx, row in enumerate(corrected_dataset.rows):
            # Row numbers are 1-based for user-facing messages
            # Add header_rows_count to get actual Excel row number
            excel_row_number = raw_dataset.header_row + raw_dataset.header_rows_count + idx + 1

            normalized_row = self.normalize_row(row, row_number=excel_row_number)

            # Track failures for this row
            if "_normalization_failures" in normalized_row:
                rows_with_failures += 1
                failed_rows.append(excel_row_number)
                total_field_failures += len(normalized_row["_normalization_failures"])

                # Log warning for row with failures
                logger.warning(
                    f"Row {excel_row_number} had {len(normalized_row['_normalization_failures'])} "
                    f"field(s) that failed normalization: {', '.join(normalized_row['_normalization_failures'])}"
                )

            normalized_rows.append(normalized_row)

        # ------------------------------------------------------------------
        # List-level one-way majority correction for birth years (web path).
        # The DateFieldProcessor applies this for the Excel-writer path; here
        # we apply the same logic to the JSON rows produced by the pipeline.
        # Only auto-completed shortened years are eligible; explicit 4-digit
        # years stored in _birth_year_auto_completed=False rows are untouched.
        # ------------------------------------------------------------------
        if self.apply_date_normalization_enabled and self.date_engine:
            normalized_rows = self._apply_birth_year_majority_correction(normalized_rows)

        # Update the rows in the dataset
        corrected_dataset.rows = normalized_rows

        # Update metadata with normalization info
        if corrected_dataset.metadata is None:
            corrected_dataset.metadata = {}

        corrected_dataset.metadata["normalized"] = True
        corrected_dataset.metadata["normalization_engines"] = {
            "name": self.apply_name_normalization_enabled and self.name_engine is not None,
            "gender": self.apply_gender_normalization_enabled and self.gender_engine is not None,
            "date": self.apply_date_normalization_enabled and self.date_engine is not None,
            "identifier": self.apply_identifier_normalization_enabled and self.identifier_engine is not None
        }
        
        # Add failure statistics to metadata
        corrected_dataset.metadata["normalization_statistics"] = {
            "total_rows": total_rows,
            "rows_with_failures": rows_with_failures,
            "total_field_failures": total_field_failures,
            "failed_rows": failed_rows,
            "success_rate": (total_rows - rows_with_failures) / total_rows if total_rows > 0 else 1.0
        }
        
        # Log summary
        if rows_with_failures > 0:
            logger.warning(
                f"Normalization completed for sheet '{raw_dataset.sheet_name}': "
                f"{rows_with_failures}/{total_rows} rows had failures "
                f"({total_field_failures} total field failures)"
            )
        else:
            logger.info(
                f"Normalization completed successfully for sheet '{raw_dataset.sheet_name}': "
                f"all {total_rows} rows processed without errors"
            )

        return corrected_dataset


    def _apply_birth_year_majority_correction(self, rows: List[JsonRow]) -> List[JsonRow]:
        """One-way list-level majority correction for birth years in the web/JSON path.

        Mirrors DateFieldProcessor._apply_majority_century_correction but operates
        on JSON rows instead of DateParseResult objects.

        Rules (identical to the Excel-writer path):
        - Only auto-completed years (tagged _birth_year_auto_completed=True) are
          considered and eligible for correction.
        - Explicit 4-digit years (tagged False) are never touched.
        - If the majority of auto-completed birth years are in the 1900s, flip
          any auto-completed 2000s years to their 1900s equivalents.
        - The reverse (flipping 1900s to 2000s) is never done.
        - After flipping, re-run validate_business_rules so status is correct.
        - The internal tag key is stripped from the final rows.
        """
        from ..data_types import DateFieldType, DateFormatPattern

        # Determine which year field to inspect: split (birth_year_corrected)
        # or single (birth_date_corrected contains DD/MM/YYYY string).
        # We use the _birth_year_auto_completed tag to identify eligible rows
        # and read the corrected year from birth_year_corrected when present,
        # otherwise parse it from birth_date_corrected.

        def _get_corrected_year(row):
            """Extract the corrected birth year integer from a row, or None."""
            if "birth_year_corrected" in row:
                try:
                    return int(row["birth_year_corrected"])
                except (TypeError, ValueError):
                    return None
            return None

        auto_1900s = sum(
            1 for r in rows
            if r.get("_birth_year_auto_completed") is True
            and _get_corrected_year(r) is not None
            and 1900 <= _get_corrected_year(r) <= 1999
        )
        auto_2000s = sum(
            1 for r in rows
            if r.get("_birth_year_auto_completed") is True
            and _get_corrected_year(r) is not None
            and 2000 <= _get_corrected_year(r) <= 2099
        )

        total_auto = auto_1900s + auto_2000s
        do_correction = total_auto > 0 and auto_1900s > auto_2000s

        corrected_rows = []
        for row in rows:
            row = dict(row)  # shallow copy so we don't mutate the original
            is_auto = row.get("_birth_year_auto_completed") is True
            yr = _get_corrected_year(row)

            if do_correction and is_auto and yr is not None and 2000 <= yr <= 2099:
                new_yr = yr - 100  # e.g. 2026 → 1926

                if "birth_year_corrected" in row:
                    # Update the corrected year and re-validate
                    mo = row.get("birth_month_corrected")
                    dy = row.get("birth_day_corrected")
                    try:
                        new_result = self.date_engine._validate_date(new_yr, mo, dy)
                        new_result.year_was_auto_completed = True
                        new_result = self.date_engine.validate_business_rules(
                            new_result, DateFieldType.BIRTH_DATE
                        )
                        row["birth_year_corrected"] = new_result.year if new_result.year is not None else new_yr
                        row["birth_month_corrected"] = new_result.month if new_result.month is not None else mo
                        row["birth_day_corrected"] = new_result.day if new_result.day is not None else dy
                        row["birth_date_status"] = new_result.status_text
                    except Exception:
                        row["birth_year_corrected"] = new_yr

            # Strip the internal tag — it must not appear in the UI payload
            row.pop("_birth_year_auto_completed", None)
            row.pop("_entry_year_auto_completed", None)
            corrected_rows.append(row)

        return corrected_rows

