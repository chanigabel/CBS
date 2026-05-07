"""Core package for Excel data standardization and export."""

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
from .orchestrator import StandardizationOrchestrator
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
    "StandardizationOrchestrator",
    "JsonExporter",
    "generate_output_filenames",
]

# Backward-compatible alias for callers that still import the legacy name.
standardizationOrchestrator = StandardizationOrchestrator
