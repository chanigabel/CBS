"""JSON export utilities for Excel standardization pipeline.

This module provides utilities to export SheetDataset and WorkbookDataset
instances to JSON files, preserving the exact structure of the data.

The pipeline produces two JSON files:
1. Raw JSON: Exact data extracted from Excel (no modifications)
2. Normalized JSON: Data with corrected fields added by standardization engines

Requirements:
    - Validates: Requirements 11.6, 19.1-19.5
"""

import json
from pathlib import Path
from typing import Any, Dict, List
from datetime import datetime, date

from .data_types import SheetDataset, WorkbookDataset, JsonRow


class JsonExporter:
    """Export datasets to JSON files.
    
    This class handles exporting SheetDataset and WorkbookDataset instances
    to JSON files with proper formatting and structure.
    
    Features:
        - Exports raw JSON (original values only)
        - Exports normalized JSON (original + corrected values)
        - Handles datetime serialization
        - Pretty-prints JSON for readability
        - Preserves metadata
    
    Example:
        exporter = JsonExporter()
        
        # Export raw dataset
        exporter.export_dataset_to_json(
            raw_dataset,
            "output_raw.json"
        )
        
        # Export normalized dataset
        exporter.export_dataset_to_json(
            normalized_dataset,
            "output_normalized.json"
        )
    """
    
    def __init__(self, indent: int = 2, ensure_ascii: bool = False):
        """Initialize JsonExporter.
        
        Args:
            indent: Number of spaces for JSON indentation (default: 2)
            ensure_ascii: If False, allow Unicode characters in output (default: False)
        """
        self.indent = indent
        self.ensure_ascii = ensure_ascii
    
    def export_dataset_to_json(self, dataset: SheetDataset, output_path: str) -> None:
        """Export a SheetDataset to JSON file.
        
        Args:
            dataset: SheetDataset to export
            output_path: Path for output JSON file
        
        Example:
            dataset = SheetDataset(
                sheet_name="Students",
                header_row=1,
                header_rows_count=1,
                field_names=["first_name", "last_name"],
                rows=[
                    {"first_name": "John", "last_name": "Doe"}
                ],
                metadata={}
            )
            
            exporter.export_dataset_to_json(dataset, "students.json")
        """
        # Convert dataset to dictionary
        dataset_dict = {
            "sheet_name": dataset.sheet_name,
            "header_row": dataset.header_row,
            "header_rows_count": dataset.header_rows_count,
            "field_names": dataset.field_names,
            "rows": self._serialize_rows(dataset.rows),
            "metadata": dataset.metadata
        }
        
        # Write to file
        self._write_json_file(dataset_dict, output_path)
    
    def export_workbook_to_json(self, workbook_dataset: WorkbookDataset, output_path: str) -> None:
        """Export a WorkbookDataset to JSON file.
        
        Args:
            workbook_dataset: WorkbookDataset to export
            output_path: Path for output JSON file
        
        Example:
            workbook_dataset = WorkbookDataset(
                source_file="data.xlsx",
                sheets=[sheet1, sheet2],
                metadata={}
            )
            
            exporter.export_workbook_to_json(workbook_dataset, "workbook.json")
        """
        # Convert workbook to dictionary
        workbook_dict = {
            "source_file": workbook_dataset.source_file,
            "sheets": [
                {
                    "sheet_name": sheet.sheet_name,
                    "header_row": sheet.header_row,
                    "header_rows_count": sheet.header_rows_count,
                    "field_names": sheet.field_names,
                    "rows": self._serialize_rows(sheet.rows),
                    "metadata": sheet.metadata
                }
                for sheet in workbook_dataset.sheets
            ],
            "metadata": workbook_dataset.metadata
        }
        
        # Write to file
        self._write_json_file(workbook_dict, output_path)
    
    def _serialize_rows(self, rows: List[JsonRow]) -> List[Dict[str, Any]]:
        """Serialize rows, handling datetime objects.
        
        Args:
            rows: List of JSON row dictionaries
        
        Returns:
            List of serialized row dictionaries
        """
        serialized_rows = []
        
        for row in rows:
            serialized_row = {}
            for key, value in row.items():
                # Convert datetime/date to ISO format string
                if isinstance(value, datetime):
                    serialized_row[key] = value.isoformat()
                elif isinstance(value, date):
                    serialized_row[key] = value.isoformat()
                else:
                    serialized_row[key] = value
            serialized_rows.append(serialized_row)
        
        return serialized_rows
    
    def _write_json_file(self, data: Dict[str, Any], output_path: str) -> None:
        """Write data to JSON file with proper formatting.
        
        Args:
            data: Dictionary to write
            output_path: Path for output file
        """
        output_file = Path(output_path)
        
        # Create parent directory if it doesn't exist
        output_file.parent.mkdir(parents=True, exist_ok=True)
        
        # Write JSON file
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(
                data,
                f,
                indent=self.indent,
                ensure_ascii=self.ensure_ascii,
                default=str  # Fallback for any non-serializable objects
            )


def generate_output_filenames(input_excel_path: str) -> tuple[str, str]:
    """Generate output filenames for raw and normalized JSON files.
    
    Given an input Excel file path, generates appropriate output filenames
    for the raw and normalized JSON files in the same directory.
    
    Args:
        input_excel_path: Path to input Excel file
    
    Returns:
        Tuple of (raw_json_path, normalized_json_path)
    
    Example:
        >>> generate_output_filenames("data/Automations_DEV.xlsx")
        ('data/Automations_DEV_raw.json', 'data/Automations_DEV_normalized.json')
    """
    input_path = Path(input_excel_path)
    
    # Get directory and base name (without extension)
    directory = input_path.parent
    base_name = input_path.stem
    
    # Generate output filenames
    raw_json_path = directory / f"{base_name}_raw.json"
    normalized_json_path = directory / f"{base_name}_normalized.json"
    
    return str(raw_json_path), str(normalized_json_path)
