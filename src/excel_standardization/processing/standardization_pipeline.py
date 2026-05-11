"""StandardizationPipeline: Apply standardization engines to JSON rows.

This module provides the StandardizationPipeline class that orchestrates the
application of standardization engines (NameEngine, GenderEngine, DateEngine,
IdentifierEngine) to JSON row data extracted from Excel worksheets.

The pipeline operates on JSON data structures, maintaining clean separation
between IO operations and business logic. It preserves original values while
creating corrected fields with the "_corrected" suffix.
"""

import logging
from datetime import date
from typing import Optional, Dict, Any, List, Tuple
from ..data_types import JsonRow, SheetDataset
from ..engines.name_engine import NameEngine
from ..engines.gender_engine import GenderEngine
from ..engines.date_engine import DateEngine
from ..engines.identifier_engine import IdentifierEngine
from ..engines.text_processor import TextProcessor
from . import date_standardization
from . import gender_standardization
from . import identifier_standardization
from . import name_standardization

# Configure logger for this module
logger = logging.getLogger(__name__)


# המחלקה משמשת מתאם מרכזי שמריץ את מנועי הסטנדרטיזציה על SheetDataset.
class StandardizationPipeline:
    """Apply standardization engines to JSON rows.
    
    This class orchestrates the application of standardization engines to JSON
    row data. It acts as an adapter between JSON data structures and the
    existing standardization engines, which operate on string and numeric values.
    
    The pipeline:
    - Accepts JSON rows with original field values
    - Applies configured standardization engines
    - Creates corrected fields with "_corrected" suffix
    - Preserves original values (non-destructive)
    - Handles missing fields and engine failures gracefully
    
    Attributes:
        name_engine: Engine for standardizing name fields
        gender_engine: Engine for standardizing gender values
        date_engine: Engine for parsing and validating dates
        identifier_engine: Engine for validating ID and passport values
        apply_name_standardization_enabled: Whether to apply name standardization
        apply_gender_standardization_enabled: Whether to apply gender standardization
        apply_date_standardization_enabled: Whether to apply date standardization
        apply_identifier_standardization_enabled: Whether to apply identifier standardization
    
    Example:
        # Create pipeline with all engines
        pipeline = StandardizationPipeline(
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
    
    # הפונקציה מקבלת את המנועים ואת דגלי ההפעלה שמגדירים אילו תיקונים ירוצו.
    def __init__(
        self,
        name_engine: Optional[NameEngine] = None,
        gender_engine: Optional[GenderEngine] = None,
        date_engine: Optional[DateEngine] = None,
        identifier_engine: Optional[IdentifierEngine] = None,
        apply_name_standardization_enabled: bool = True,
        apply_gender_standardization_enabled: bool = True,
        apply_date_standardization_enabled: bool = True,
        apply_identifier_standardization_enabled: bool = True,
        reference_date: Optional[date] = None,
    ):
        """Initialize StandardizationPipeline with engine dependencies.
        
        Args:
            name_engine: Engine for standardizing name fields (optional)
            gender_engine: Engine for standardizing gender values (optional)
            date_engine: Engine for parsing and validating dates (optional)
            identifier_engine: Engine for validating ID and passport values (optional)
            apply_name_standardization_enabled: Whether to apply name standardization
            apply_gender_standardization_enabled: Whether to apply gender standardization
            apply_date_standardization_enabled: Whether to apply date standardization
            apply_identifier_standardization_enabled: Whether to apply identifier standardization
        
        Note:
            If an engine is not provided, the corresponding standardization will be
            skipped even if the enabled flag is True. This allows for flexible
            configuration where only specific engines are used.
        
        Requirements:
            - Validates: Requirements 12.1, 17.6
        """
        self.name_engine = name_engine
        self.gender_engine = gender_engine
        self.date_engine = date_engine
        self.identifier_engine = identifier_engine
        self._reference_date = reference_date or date.today()
        if self.date_engine is not None:
            self.date_engine.reference_date = self._reference_date
        
        # Configuration flags for which engines to apply
        self.apply_name_standardization_enabled = apply_name_standardization_enabled
        self.apply_gender_standardization_enabled = apply_gender_standardization_enabled
        self.apply_date_standardization_enabled = apply_date_standardization_enabled
        self.apply_identifier_standardization_enabled = apply_identifier_standardization_enabled
    
    # הפונקציה מנרמלת שורה אחת ומוסיפה שדות corrected בלי לשנות את ערכי המקור.
    def normalize_row(self, json_row: JsonRow, row_number: Optional[int] = None) -> JsonRow:
        """Apply standardization engines to a single row.
        
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
        
        # Track failed standardizations for this row
        failed_fields: List[str] = []
        
        # Apply each standardization engine
        if self.apply_name_standardization_enabled and self.name_engine:
            failures = self.apply_name_standardization(result, row_number)
            failed_fields.extend(failures)
        
        if self.apply_gender_standardization_enabled and self.gender_engine:
            failures = self.apply_gender_standardization(result, row_number)
            failed_fields.extend(failures)
        
        if self.apply_date_standardization_enabled and self.date_engine:
            failures = self.apply_date_standardization(result, row_number)
            failed_fields.extend(failures)
        
        if self.apply_identifier_standardization_enabled and self.identifier_engine:
            failures = self.apply_identifier_standardization(result, row_number)
            failed_fields.extend(failures)
        
        # Store failed fields in metadata if any failures occurred
        if failed_fields:
            result["_standardization_failures"] = failed_fields
        
        return result
    
    # הפונקציה מפעילה תיקוני שמות על השורה דרך מודול name_standardization.
    def apply_name_standardization(self, json_row: JsonRow, row_number: Optional[int] = None) -> List[str]:
        """Apply NameEngine to name fields in the row.

        Updates json_row with corrected fields for:
        - first_name  -> first_name_corrected  (with last-name removal if applicable)
        - last_name   -> last_name_corrected
        - father_name -> father_name_corrected (with last-name removal if applicable)

        Last-name removal uses the two-stage logic in NameEngine:
          Stage A: substring removal.
          Stage B: positional fallback ? only when Stage A made no change.

        The pattern for father_name and first_name is stored on the pipeline
        instance (set by normalize_dataset before iterating rows).

        Args:
            json_row: Dictionary to update with corrected name fields
            row_number: Optional row number for error logging (1-based)

        Returns:
            List of field names that failed standardization

        Requirements:
            - Validates: Requirements 12.3, 12.8, 14.1-14.5, 18.1-18.4
        """
        return name_standardization.apply_name_standardization(self, json_row, row_number)

    # הפונקציה מפעילה תיקון מגדר ומעדכנת gender_corrected/status בשורת ה־Dataset.
    def apply_gender_standardization(self, json_row: JsonRow, row_number: Optional[int] = None) -> List[str]:
        """Apply GenderEngine to gender field in the row.
        
        Updates json_row with corrected field:
        - gender -> gender_corrected
        
        Args:
            json_row: Dictionary to update with corrected gender field
            row_number: Optional row number for error logging (1-based)
        
        Returns:
            List of field names that failed standardization
        
        Requirements:
            - Validates: Requirements 12.4, 12.8, 14.1-14.5, 18.1-18.4
        """
        return gender_standardization.apply_gender_standardization(self, json_row, row_number)

    # הפונקציה מפעילה תיקוני תאריך לידה וכניסה ומוסיפה שדות corrected/status.
    def apply_date_standardization(self, json_row: JsonRow, row_number: Optional[int] = None) -> List[str]:
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
            List of field names that failed standardization
        
        Requirements:
            - Validates: Requirements 12.5, 12.8, 14.1-14.5, 18.1-18.4
        """
        return date_standardization.apply_date_standardization(self, json_row, row_number)
    
    # הפונקציה מנרמלת שדה תאריך יחיד או מפוצל עבור prefix נתון.
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
        return date_standardization.normalize_date_field(self, json_row, prefix, field_type, row_number)

    # הפונקציה מחזירה רכיבי תאריך בטוחים להצגה וליצוא לאחר parsing.
    def _date_corrected_components(self, result) -> Tuple[Any, Any, Any]:
        """Return UI/export-safe corrected date components.

        Invalid raw split-date values must never be copied into corrected
        fields. Component-level range errors blank only the failing component;
        non-numeric date content blanks any component that could not be parsed.
        """
        return date_standardization.date_corrected_components(result)
    
    # הפונקציה מפעילה תיקוני תעודת זהות ודרכון ומעדכנת שדות corrected.
    def apply_identifier_standardization(self, json_row: JsonRow, row_number: Optional[int] = None) -> List[str]:
        """Apply IdentifierEngine to identifier fields in the row.
        
        Updates json_row with corrected fields for:
        - id_number -> id_number_corrected
        - passport -> passport_corrected
        
        Args:
            json_row: Dictionary to update with corrected identifier fields
            row_number: Optional row number for error logging (1-based)
        
        Returns:
            List of field names that failed standardization
        
        Requirements:
            - Validates: Requirements 12.6, 12.8, 14.1-14.5, 18.1-18.4
        """
        return identifier_standardization.apply_identifier_standardization(self, json_row, row_number)

    # הפונקציה מנרמלת גיליון שלם, מחשבת סטטיסטיקות ומריצה validation לאחר התיקונים.
    def normalize_dataset(self, raw_dataset: SheetDataset) -> SheetDataset:
        """Apply standardization engines to all rows in dataset.

        Creates a new dataset with both original and corrected values.
        Updates metadata with standardization information and tracks failed standardizations.

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
        if self.apply_name_standardization_enabled and self.name_engine:
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
            # apply_name_standardization can read them per-row.
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
        if self.apply_date_standardization_enabled and self.date_engine:
            self._date_format_pattern = date_standardization.detect_date_format_pattern(corrected_dataset.rows)
            logger.debug(
                f"Date format pattern detected for sheet '{raw_dataset.sheet_name}': "
                f"{self._date_format_pattern}"
            )
        else:
            from ..data_types import DateFormatPattern
            self._date_format_pattern = DateFormatPattern.DDMM

        # Track standardization statistics
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
            if "_standardization_failures" in normalized_row:
                rows_with_failures += 1
                failed_rows.append(excel_row_number)
                total_field_failures += len(normalized_row["_standardization_failures"])

                # Log warning for row with failures
                logger.warning(
                    f"Row {excel_row_number} had {len(normalized_row['_standardization_failures'])} "
                    f"field(s) that failed standardization: {', '.join(normalized_row['_standardization_failures'])}"
                )

            normalized_rows.append(normalized_row)

        # ------------------------------------------------------------------
        # List-level one-way majority correction for birth years (web path).
        # The DateFieldProcessor applies this for the Excel-writer path; here
        # we apply the same logic to the JSON rows produced by the pipeline.
        # Only auto-completed shortened years are eligible; explicit 4-digit
        # years stored in _birth_year_auto_completed=False rows are untouched.
        # ------------------------------------------------------------------
        if self.apply_date_standardization_enabled and self.date_engine:
            normalized_rows = self._apply_birth_year_majority_correction(normalized_rows)

        if (
            self.apply_identifier_standardization_enabled
            and any(row.get("passport_corrected") for row in normalized_rows)
            and "passport_corrected" not in corrected_dataset.field_names
        ):
            corrected_dataset.field_names = list(corrected_dataset.field_names)
            corrected_dataset.field_names.append("passport_corrected")

        # Update the rows in the dataset
        corrected_dataset.rows = normalized_rows

        # ------------------------------------------------------------------
        # Institution-report validation (post-normalization).
        # Runs after all corrected fields are written so validators can read
        # *_corrected values.  Results are written into each row dict under
        # _validation_status and _validation_ok.
        # Only runs for the three known institution-report sheet types.
        # ------------------------------------------------------------------
        try:
            from ..validation.institution_report_validator import (
                InstitutionReportValidator,
                KNOWN_SHEETS,
            )
            from ..services.sheet_name_resolver import resolve_canonical_sheet_name

            canonical = resolve_canonical_sheet_name(corrected_dataset.sheet_name)
            if canonical in KNOWN_SHEETS:
                # Pass MosadID from sheet metadata if available (set by mosad_id_scanner).
                sheet_mosad_id = corrected_dataset.get_metadata("MosadID")
                validator = InstitutionReportValidator(sheet_name=canonical)
                validator.validate_sheet(
                    corrected_dataset.rows,
                    sheet_mosad_id=sheet_mosad_id,
                )
                logger.debug(
                    "Institution-report validation completed for sheet '%s' (canonical: '%s')",
                    corrected_dataset.sheet_name,
                    canonical,
                )
        except Exception as _val_exc:
            logger.warning(
                "Institution-report validation skipped for sheet '%s': %s",
                corrected_dataset.sheet_name,
                _val_exc,
            )

        # Update metadata with standardization info
        if corrected_dataset.metadata is None:
            corrected_dataset.metadata = {}

        corrected_dataset.metadata["normalized"] = True
        corrected_dataset.metadata["standardization_engines"] = {
            "name": self.apply_name_standardization_enabled and self.name_engine is not None,
            "gender": self.apply_gender_standardization_enabled and self.gender_engine is not None,
            "date": self.apply_date_standardization_enabled and self.date_engine is not None,
            "identifier": self.apply_identifier_standardization_enabled and self.identifier_engine is not None
        }
        corrected_dataset.metadata["processing_date"] = self._reference_date.isoformat()
        corrected_dataset.metadata["processing_year"] = self._reference_date.year
        
        # Add failure statistics to metadata
        corrected_dataset.metadata["standardization_statistics"] = {
            "total_rows": total_rows,
            "rows_with_failures": rows_with_failures,
            "total_field_failures": total_field_failures,
            "failed_rows": failed_rows,
            "success_rate": (total_rows - rows_with_failures) / total_rows if total_rows > 0 else 1.0
        }
        
        # Log summary
        if rows_with_failures > 0:
            logger.warning(
                f"standardization completed for sheet '{raw_dataset.sheet_name}': "
                f"{rows_with_failures}/{total_rows} rows had failures "
                f"({total_field_failures} total field failures)"
            )
        else:
            logger.info(
                f"standardization completed successfully for sheet '{raw_dataset.sheet_name}': "
                f"all {total_rows} rows processed without errors"
            )

        return corrected_dataset

    # הפונקציה מחילה תיקון רוב לשנת לידה מקוצרת אחרי נרמול כל השורות.
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
        return date_standardization.apply_birth_year_majority_correction(self, rows)

# Backward-compatible alias for callers that still import the legacy name.
standardizationPipeline = StandardizationPipeline

