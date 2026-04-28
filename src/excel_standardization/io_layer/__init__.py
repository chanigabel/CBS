"""I/O layer for Excel operations.

This package contains ExcelReader and ExcelWriter classes that encapsulate
all openpyxl interactions, isolating Excel I/O from business logic.
"""

from .excel_reader import ExcelReader
from .excel_to_json_extractor import ExcelToJsonExtractor
from .excel_writer import ExcelWriter, JsonToExcelWriter

__all__ = ["ExcelReader", "ExcelToJsonExtractor", "ExcelWriter", "JsonToExcelWriter"]
