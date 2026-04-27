"""Excel Data Normalization System.

A Python-based system that replicates the exact behavior of a legacy VBA
implementation for normalizing Excel data containing person records.
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
from .orchestrator import NormalizationOrchestrator
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
    "NormalizationOrchestrator",
    "JsonExporter",
    "generate_output_filenames",
]
