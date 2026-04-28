"""Core data types and enums for the Excel standardization system.

This module defines all dataclasses and enums used throughout the system
to represent processing results and configuration options.
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional


# ============================================================================
# JSON Data Types
# ============================================================================

JsonRow = Dict[str, Any]
"""Type alias for a JSON row representing a single data row from Excel.

A JsonRow is a dictionary with field names as keys and cell values as values.
It follows a specific naming convention for original and corrected fields:

Field Naming Convention:
    - Original fields: Use standardized field names (e.g., "first_name", "gender")
    - Corrected fields: Use "_corrected" suffix (e.g., "first_name_corrected", "gender_corrected")

Common Field Types:
    - Name fields (str): first_name, last_name, father_name
    - Gender field (str|int): gender (original may be Hebrew text or number)
    - Identifier fields (str): id_number, passport
    - Date fields (str|int): 
        * Single date: birth_date, entry_date (str)
        * Split date: birth_year, birth_month, birth_day (int)
        * Split date: entry_year, entry_month, entry_day (int)

Example:
    {
        "first_name": "יוסי",
        "first_name_corrected": "יוסי",
        "last_name": "כהן",
        "last_name_corrected": "כהן",
        "gender": "ז",
        "gender_corrected": "2",
        "id_number": "123456789",
        "id_number_corrected": "123456789",
        "birth_year": 1980,
        "birth_year_corrected": 1980,
        "birth_month": 5,
        "birth_month_corrected": 5,
        "birth_day": 15,
        "birth_day_corrected": 15
    }

Requirements:
    - Validates: Requirements 10.2, 14.3-14.4
"""


@dataclass
class SheetDataset:
    """Dataset for a single worksheet containing extracted JSON rows and metadata.
    
    This dataclass represents all data extracted from a single Excel worksheet,
    including the JSON rows, header information, and metadata about the sheet structure.
    
    Attributes:
        sheet_name: Name of the worksheet
        header_row: Row number where headers were found (1-based)
        header_rows_count: Number of header rows (1 or 2 for multi-row headers)
        field_names: List of detected field names in order
        rows: List of JSON row dictionaries containing the data
        metadata: Additional metadata about the sheet (optional)
    
    Metadata Dictionary Keys:
        - source_file: Path to source Excel file
        - extraction_date: Date when data was extracted
        - total_rows: Total number of data rows
        - date_field_structure: Dict mapping date fields to "single" or "split"
        - skipped_rows: Number of rows skipped during extraction
        - errors: List of error messages encountered during extraction
    
    Example:
        dataset = SheetDataset(
            sheet_name="Students",
            header_row=2,
            header_rows_count=1,
            field_names=["first_name", "last_name", "gender", "birth_year"],
            rows=[
                {"first_name": "יוסי", "last_name": "כהן", "gender": "ז", "birth_year": 1980},
                {"first_name": "שרה", "last_name": "לוי", "gender": "נ", "birth_year": 1985}
            ],
            metadata={
                "source_file": "data.xlsx",
                "total_rows": 2,
                "date_field_structure": {"birth_date": "split"}
            }
        )
    
    Requirements:
        - Validates: Requirements 11.2-11.5, 19.1-19.2
    """
    
    sheet_name: str
    header_row: int
    header_rows_count: int
    field_names: List[str]
    rows: List[JsonRow]
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def get_field_names(self) -> List[str]:
        """Get the list of field names for this dataset.
        
        Returns:
            List of field names in order
        """
        return self.field_names
    
    def get_row_count(self) -> int:
        """Get the number of data rows in this dataset.
        
        Returns:
            Number of rows
        """
        return len(self.rows)
    
    def validate(self) -> bool:
        """Validate the dataset structure.
        
        Checks:
            - sheet_name is not empty
            - header_row is positive
            - header_rows_count is 1 or 2
            - field_names is not empty
            - rows is a list
            - All rows are dictionaries
        
        Returns:
            True if valid, False otherwise
        """
        if not self.sheet_name:
            return False
        
        if self.header_row < 1:
            return False
        
        if self.header_rows_count not in (1, 2):
            return False
        
        if not self.field_names:
            return False
        
        if not isinstance(self.rows, list):
            return False
        
        for row in self.rows:
            if not isinstance(row, dict):
                return False
        
        return True
    
    def get_metadata(self, key: str, default: Any = None) -> Any:
        """Get a metadata value by key.
        
        Args:
            key: Metadata key to retrieve
            default: Default value if key not found
        
        Returns:
            Metadata value or default
        """
        return self.metadata.get(key, default)
    
    def set_metadata(self, key: str, value: Any) -> None:
        """Set a metadata value.
        
        Args:
            key: Metadata key to set
            value: Value to store
        """
        self.metadata[key] = value


@dataclass
class WorkbookDataset:
    """Dataset for an entire workbook containing multiple SheetDataset instances.
    
    This dataclass represents all data extracted from an Excel workbook,
    including multiple worksheets, each with their own JSON rows and metadata.
    
    Attributes:
        source_file: Path to the source Excel file
        sheets: List of SheetDataset instances, one per worksheet
        metadata: Workbook-level metadata (optional)
    
    Metadata Dictionary Keys:
        - extraction_date: Date when data was extracted
        - total_sheets: Total number of sheets in workbook
        - processed_sheets: Number of sheets successfully processed
        - skipped_sheets: List of sheet names that were skipped
        - errors: List of error messages encountered during extraction
    
    Example:
        dataset = WorkbookDataset(
            source_file="data.xlsx",
            sheets=[
                SheetDataset(
                    sheet_name="Students",
                    header_row=2,
                    header_rows_count=1,
                    field_names=["first_name", "last_name"],
                    rows=[{"first_name": "יוסי", "last_name": "כהן"}],
                    metadata={}
                ),
                SheetDataset(
                    sheet_name="Teachers",
                    header_row=1,
                    header_rows_count=1,
                    field_names=["first_name", "last_name"],
                    rows=[{"first_name": "שרה", "last_name": "לוי"}],
                    metadata={}
                )
            ],
            metadata={
                "extraction_date": "2024-01-15",
                "total_sheets": 3,
                "processed_sheets": 2,
                "skipped_sheets": ["Summary"]
            }
        )
    
    Requirements:
        - Validates: Requirements 16.1-16.4
    """
    
    source_file: str
    sheets: List[SheetDataset]
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def get_sheet_by_name(self, sheet_name: str) -> Optional[SheetDataset]:
        """Get a sheet dataset by name.
        
        Args:
            sheet_name: Name of the sheet to retrieve
        
        Returns:
            SheetDataset if found, None otherwise
        """
        for sheet in self.sheets:
            if sheet.sheet_name == sheet_name:
                return sheet
        return None
    
    def get_sheet_names(self) -> List[str]:
        """Get list of all sheet names in the workbook.
        
        Returns:
            List of sheet names in order
        """
        return [sheet.sheet_name for sheet in self.sheets]
    
    def get_sheet_count(self) -> int:
        """Get the number of sheets in this workbook dataset.
        
        Returns:
            Number of sheets
        """
        return len(self.sheets)
    
    def validate(self) -> bool:
        """Validate the workbook dataset structure.
        
        Checks:
            - source_file is not empty
            - sheets is a list
            - All sheets are SheetDataset instances
            - All sheets have unique names
            - All sheets are valid
        
        Returns:
            True if valid, False otherwise
        """
        if not self.source_file:
            return False
        
        if not isinstance(self.sheets, list):
            return False
        
        # Check all sheets are SheetDataset instances
        for sheet in self.sheets:
            if not isinstance(sheet, SheetDataset):
                return False
        
        # Check all sheets have unique names
        sheet_names = [sheet.sheet_name for sheet in self.sheets]
        if len(sheet_names) != len(set(sheet_names)):
            return False
        
        # Validate each sheet
        for sheet in self.sheets:
            if not sheet.validate():
                return False
        
        return True
    
    def get_metadata(self, key: str, default: Any = None) -> Any:
        """Get a metadata value by key.
        
        Args:
            key: Metadata key to retrieve
            default: Default value if key not found
        
        Returns:
            Metadata value or default
        """
        return self.metadata.get(key, default)
    
    def set_metadata(self, key: str, value: Any) -> None:
        """Set a metadata value.
        
        Args:
            key: Metadata key to set
            value: Value to store
        """
        self.metadata[key] = value
    
    def has_sheet(self, sheet_name: str) -> bool:
        """Check if a sheet with the given name exists.
        
        Args:
            sheet_name: Name of the sheet to check
        
        Returns:
            True if sheet exists, False otherwise
        """
        return self.get_sheet_by_name(sheet_name) is not None


# ============================================================================
# Excel Data Types
# ============================================================================


@dataclass
class ColumnHeaderInfo:
    """Information about a found column header.

    Attributes:
        col: Column number (1-based)
        header_row: Row number where header was found
        last_row: Last row with data in this column
        header_text: The actual header text found
    """

    col: int
    header_row: int
    last_row: int
    header_text: str


@dataclass
class TableRegion:
    """Information about a detected table region in a worksheet.

    Attributes:
        start_row: First row of the table (header row)
        end_row: Last row of the table (data)
        start_col: First column of the table
        end_col: Last column of the table
        header_rows: Number of header rows (1 or 2 for grouped headers)
        data_start_row: First row containing actual data
    """

    start_row: int
    end_row: int
    start_col: int
    end_col: int
    header_rows: int
    data_start_row: int


@dataclass
class DateGroup:
    """Deterministic description of a split date group (Year/Month/Day).

    Used by the JSON extraction pipeline to ensure birth/entry split date columns
    are detected and grouped consistently (VBA parity intent).
    """

    year_col: int
    month_col: int
    day_col: int
    main_col: Optional[int]
    field_type: "DateFieldType"


@dataclass
class DateParseResult:
    """Result of date parsing operation.

    Attributes:
        year: Parsed year value (None if parsing failed)
        month: Parsed month value (None if parsing failed)
        day: Parsed day value (None if parsing failed)
        is_valid: Whether the date is valid
        status_text: Hebrew status message describing the result
        year_was_auto_completed: True when the year was expanded from a
            shortened (1- or 2-digit) input; False for explicit 4-digit years.
            Used by the list-level majority correction to avoid touching
            explicitly written years.
    """

    year: Optional[int]
    month: Optional[int]
    day: Optional[int]
    is_valid: bool
    status_text: str
    year_was_auto_completed: bool = False


@dataclass
class IdentifierResult:
    """Result of identifier processing.

    Attributes:
        corrected_id: Corrected Israeli ID value
        corrected_passport: Corrected passport value
        status_text: Hebrew status message describing the result
    """

    corrected_id: str
    corrected_passport: str
    status_text: str


class Language(Enum):
    """Language dominance in text."""

    HEBREW = "hebrew"
    ENGLISH = "english"
    MIXED = "mixed"


class FatherNamePattern(Enum):
    """Pattern for removing last name from father name."""

    NONE = "none"
    REMOVE_FIRST = "remove_first"
    REMOVE_LAST = "remove_last"


class DateFormatPattern(Enum):
    """Date format pattern for ambiguous dates."""

    DDMM = "ddmm"
    MMDD = "mmdd"


class DateFieldType(Enum):
    """Type of date field being processed."""

    BIRTH_DATE = "birth_date"
    ENTRY_DATE = "entry_date"


class FieldKey(Enum):
    """Keys for tracking corrected columns."""

    SHEM_PRATI = "ShemPrati"
    SHEM_MISHPAHA = "ShemMishpaha"
    SHEM_HAAV = "ShemHaAv"
    MIN = "Min"
    SHNAT_LIDA = "ShnatLida"
    HODESH_LIDA = "HodeshLida"
    YOM_LIDA = "YomLida"
    SHNAT_KNISA = "shnatknisa"
    HODESH_KNISA = "Hodeshknisa"
    YOM_KNISA = "YomKnisa"
    MISPAR_ZEHUT = "MisparZehut"
    DARKON = "Darkon"
