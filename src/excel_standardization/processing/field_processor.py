"""Abstract base class for field processors using Template Method pattern.

This module defines the FieldProcessor abstract base class that provides
a template method for processing different field types. Subclasses implement
the three abstract methods for specific field types (names, gender, dates, identifiers).
"""

from abc import ABC, abstractmethod
from openpyxl.worksheet.worksheet import Worksheet
from ..io_layer.excel_reader import ExcelReader
from ..io_layer.excel_writer import ExcelWriter


class FieldProcessor(ABC):
    """Abstract base class for field processors.

    This class uses the Template Method pattern. The process_field method
    is the template that calls find_headers, prepare_output_columns, and
    process_data in sequence. Subclasses implement the three abstract methods
    for specific field types.

    Attributes:
        reader: ExcelReader instance for reading data from worksheets
        writer: ExcelWriter instance for writing data to worksheets
    """

    def __init__(self, reader: ExcelReader, writer: ExcelWriter):
        """Initialize the field processor.

        Args:
            reader: ExcelReader instance for reading data
            writer: ExcelWriter instance for writing data
        """
        self.reader = reader
        self.writer = writer

    @abstractmethod
    def find_headers(self, worksheet: Worksheet) -> bool:
        """Find column headers for this field type.

        This method searches for the relevant column headers in the worksheet
        and stores the header information for later use in processing.

        Args:
            worksheet: The worksheet to search for headers

        Returns:
            True if all required headers were found, False otherwise
        """
        pass

    @abstractmethod
    def prepare_output_columns(self, worksheet: Worksheet) -> None:
        """Insert corrected columns after original columns.

        This method inserts new columns in the worksheet for the corrected
        values, typically immediately after the original columns. It also
        sets the header text for the new columns.

        Args:
            worksheet: The worksheet to modify
        """
        pass

    @abstractmethod
    def process_data(self, worksheet: Worksheet) -> None:
        """Read, normalize, and write data.

        This method reads the original data from the worksheet, applies
        standardization logic (typically by delegating to an Engine class),
        and writes the corrected values back to the worksheet. It may also
        apply formatting such as highlighting changed cells.

        Args:
            worksheet: The worksheet to process
        """
        pass

    def process_field(self, worksheet: Worksheet) -> None:
        """Template method: find headers, prepare columns, process data.

        This is the main entry point for processing a field. It orchestrates
        the three-step process:
        1. Find the relevant headers in the worksheet
        2. Prepare output columns for corrected values
        3. Process the data (read, normalize, write)

        If headers are not found (step 1 returns False), the processing
        is skipped for this field.

        Args:
            worksheet: The worksheet to process
        """
        # Step 1: Find headers
        if not self.find_headers(worksheet):
            # Headers not found, skip processing for this field
            return

        # Step 2: Prepare output columns
        self.prepare_output_columns(worksheet)

        # Step 3: Process data
        self.process_data(worksheet)
