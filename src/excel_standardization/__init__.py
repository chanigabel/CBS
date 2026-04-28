"""Excel Data standardization System.

A Python-based system that replicates the exact behavior of a legacy VBA
implementation for standardizing Excel data containing person records.
"""

__version__ = "1.0.0"

from .data_types import (
    ColumnHeaderInfo,
    DateFieldType,
    DateFormatPattern,
    DateParseResult,
    FatherNamePattern,
    FieldKey,
    IdentifierResult,
    JsonRow,
    Language,
    SheetDataset,
    TableRegion,
    WorkbookDataset,
)
from .orchestrator import standardizationOrchestrator
from .json_exporter import JsonExporter, generate_output_filenames

__all__ = [
    # Data types
    "ColumnHeaderInfo",
    "DateFieldType",
    "DateFormatPattern",
    "DateParseResult",
    "FatherNamePattern",
    "FieldKey",
    "IdentifierResult",
    "JsonRow",
    "Language",
    "SheetDataset",
    "TableRegion",
    "WorkbookDataset",
    # Main entry points
    "standardizationOrchestrator",
    "JsonExporter",
    "generate_output_filenames",
]
