"""JSON schema validation utilities for Excel standardization data structures.

This module provides utilities to validate SheetDataset, WorkbookDataset, and JsonRow
instances against their JSON schemas. It uses the jsonschema library for validation.

Requirements:
    - Validates: Requirements 19.1-19.5
"""

import json
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    import jsonschema
    from jsonschema import Draft7Validator, ValidationError
    JSONSCHEMA_AVAILABLE = True
except ImportError:
    JSONSCHEMA_AVAILABLE = False

from .data_types import SheetDataset, WorkbookDataset, JsonRow


# ============================================================================
# Schema Loading
# ============================================================================

def _get_schema_path(schema_name: str) -> Path:
    """Get the path to a schema file.
    
    Args:
        schema_name: Name of the schema file (e.g., "sheet_dataset.schema.json")
    
    Returns:
        Path to the schema file
    """
    # Get the project root (parent of src directory)
    current_file = Path(__file__)
    project_root = current_file.parent.parent.parent
    schema_path = project_root / "schemas" / schema_name
    return schema_path


def load_schema(schema_name: str) -> Dict[str, Any]:
    """Load a JSON schema from file.
    
    Args:
        schema_name: Name of the schema file (e.g., "sheet_dataset.schema.json")
    
    Returns:
        Parsed JSON schema as dictionary
    
    Raises:
        FileNotFoundError: If schema file doesn't exist
        json.JSONDecodeError: If schema file is not valid JSON
    """
    schema_path = _get_schema_path(schema_name)
    
    if not schema_path.exists():
        raise FileNotFoundError(f"Schema file not found: {schema_path}")
    
    with open(schema_path, 'r', encoding='utf-8') as f:
        return json.load(f)


# ============================================================================
# Validation Functions
# ============================================================================

def validate_json_row(row: JsonRow, raise_on_error: bool = False) -> tuple[bool, Optional[List[str]]]:
    """Validate a JsonRow against the JSON schema.
    
    Args:
        row: JsonRow dictionary to validate
        raise_on_error: If True, raise ValidationError on validation failure
    
    Returns:
        Tuple of (is_valid, error_messages)
        - is_valid: True if validation passed, False otherwise
        - error_messages: List of validation error messages, None if valid
    
    Raises:
        ValidationError: If raise_on_error is True and validation fails
        ImportError: If jsonschema library is not installed
    
    Example:
        >>> row = {"first_name": "John", "first_name_corrected": "John"}
        >>> is_valid, errors = validate_json_row(row)
        >>> if not is_valid:
        ...     print(f"Validation errors: {errors}")
    """
    if not JSONSCHEMA_AVAILABLE:
        raise ImportError(
            "jsonschema library is required for schema validation. "
            "Install it with: pip install jsonschema"
        )
    
    try:
        schema = load_schema("json_row.schema.json")
        validator = Draft7Validator(schema)
        
        # Collect all validation errors
        errors = list(validator.iter_errors(row))
        
        if errors:
            if raise_on_error:
                raise ValidationError(f"JsonRow validation failed: {errors[0].message}")
            
            error_messages = [
                f"{'.'.join(str(p) for p in error.path)}: {error.message}"
                for error in errors
            ]
            return False, error_messages
        
        return True, None
    
    except (FileNotFoundError, json.JSONDecodeError) as e:
        if raise_on_error:
            raise
        return False, [str(e)]


def validate_sheet_dataset(dataset: SheetDataset, raise_on_error: bool = False) -> tuple[bool, Optional[List[str]]]:
    """Validate a SheetDataset against the JSON schema.
    
    Args:
        dataset: SheetDataset instance to validate
        raise_on_error: If True, raise ValidationError on validation failure
    
    Returns:
        Tuple of (is_valid, error_messages)
        - is_valid: True if validation passed, False otherwise
        - error_messages: List of validation error messages, None if valid
    
    Raises:
        ValidationError: If raise_on_error is True and validation fails
        ImportError: If jsonschema library is not installed
    
    Example:
        >>> dataset = SheetDataset(
        ...     sheet_name="Students",
        ...     header_row=1,
        ...     header_rows_count=1,
        ...     field_names=["first_name"],
        ...     rows=[{"first_name": "John"}],
        ...     metadata={}
        ... )
        >>> is_valid, errors = validate_sheet_dataset(dataset)
        >>> assert is_valid
    """
    if not JSONSCHEMA_AVAILABLE:
        raise ImportError(
            "jsonschema library is required for schema validation. "
            "Install it with: pip install jsonschema"
        )
    
    try:
        schema = load_schema("sheet_dataset.schema.json")
        
        # Convert dataclass to dictionary for validation
        dataset_dict = {
            "sheet_name": dataset.sheet_name,
            "header_row": dataset.header_row,
            "header_rows_count": dataset.header_rows_count,
            "field_names": dataset.field_names,
            "rows": dataset.rows,
            "metadata": dataset.metadata
        }
        
        validator = Draft7Validator(schema)
        
        # Collect all validation errors
        errors = list(validator.iter_errors(dataset_dict))
        
        if errors:
            if raise_on_error:
                raise ValidationError(f"SheetDataset validation failed: {errors[0].message}")
            
            error_messages = [
                f"{'.'.join(str(p) for p in error.path)}: {error.message}"
                for error in errors
            ]
            return False, error_messages
        
        return True, None
    
    except (FileNotFoundError, json.JSONDecodeError) as e:
        if raise_on_error:
            raise
        return False, [str(e)]


def validate_workbook_dataset(dataset: WorkbookDataset, raise_on_error: bool = False) -> tuple[bool, Optional[List[str]]]:
    """Validate a WorkbookDataset against the JSON schema.
    
    Args:
        dataset: WorkbookDataset instance to validate
        raise_on_error: If True, raise ValidationError on validation failure
    
    Returns:
        Tuple of (is_valid, error_messages)
        - is_valid: True if validation passed, False otherwise
        - error_messages: List of validation error messages, None if valid
    
    Raises:
        ValidationError: If raise_on_error is True and validation fails
        ImportError: If jsonschema library is not installed
    
    Example:
        >>> dataset = WorkbookDataset(
        ...     source_file="data.xlsx",
        ...     sheets=[],
        ...     metadata={}
        ... )
        >>> is_valid, errors = validate_workbook_dataset(dataset)
        >>> assert is_valid
    """
    if not JSONSCHEMA_AVAILABLE:
        raise ImportError(
            "jsonschema library is required for schema validation. "
            "Install it with: pip install jsonschema"
        )
    
    try:
        schema = load_schema("workbook_dataset.schema.json")
        
        # Convert dataclass to dictionary for validation
        sheets_dict = []
        for sheet in dataset.sheets:
            sheets_dict.append({
                "sheet_name": sheet.sheet_name,
                "header_row": sheet.header_row,
                "header_rows_count": sheet.header_rows_count,
                "field_names": sheet.field_names,
                "rows": sheet.rows,
                "metadata": sheet.metadata
            })
        
        dataset_dict = {
            "source_file": dataset.source_file,
            "sheets": sheets_dict,
            "metadata": dataset.metadata
        }
        
        validator = Draft7Validator(schema)
        
        # Collect all validation errors
        errors = list(validator.iter_errors(dataset_dict))
        
        if errors:
            if raise_on_error:
                raise ValidationError(f"WorkbookDataset validation failed: {errors[0].message}")
            
            error_messages = [
                f"{'.'.join(str(p) for p in error.path)}: {error.message}"
                for error in errors
            ]
            return False, error_messages
        
        return True, None
    
    except (FileNotFoundError, json.JSONDecodeError) as e:
        if raise_on_error:
            raise
        return False, [str(e)]


# ============================================================================
# Field Naming Convention Utilities
# ============================================================================

def get_corrected_field_name(field_name: str) -> str:
    """Get the corrected field name for a given original field name.
    
    Follows the naming convention: field_name → field_name_corrected
    
    Args:
        field_name: Original field name (e.g., "first_name")
    
    Returns:
        Corrected field name (e.g., "first_name_corrected")
    
    Example:
        >>> get_corrected_field_name("first_name")
        'first_name_corrected'
        >>> get_corrected_field_name("gender")
        'gender_corrected'
    
    Requirements:
        - Validates: Requirements 13.3, 14.4, 19.4
    """
    if field_name.endswith("_corrected"):
        return field_name
    return f"{field_name}_corrected"


def get_original_field_name(corrected_field_name: str) -> str:
    """Get the original field name from a corrected field name.
    
    Follows the naming convention: field_name_corrected → field_name
    
    Args:
        corrected_field_name: Corrected field name (e.g., "first_name_corrected")
    
    Returns:
        Original field name (e.g., "first_name")
    
    Example:
        >>> get_original_field_name("first_name_corrected")
        'first_name'
        >>> get_original_field_name("gender")
        'gender'
    
    Requirements:
        - Validates: Requirements 13.3, 14.3, 19.4
    """
    if corrected_field_name.endswith("_corrected"):
        return corrected_field_name[:-len("_corrected")]
    return corrected_field_name


def is_corrected_field(field_name: str) -> bool:
    """Check if a field name is a corrected field.
    
    Args:
        field_name: Field name to check
    
    Returns:
        True if field name ends with "_corrected", False otherwise
    
    Example:
        >>> is_corrected_field("first_name_corrected")
        True
        >>> is_corrected_field("first_name")
        False
    
    Requirements:
        - Validates: Requirements 13.3, 19.4
    """
    return field_name.endswith("_corrected")


def get_field_pairs(json_row: JsonRow) -> List[tuple[str, str]]:
    """Get pairs of (original_field, corrected_field) from a JsonRow.
    
    Args:
        json_row: JsonRow dictionary
    
    Returns:
        List of tuples containing (original_field_name, corrected_field_name)
    
    Example:
        >>> row = {
        ...     "first_name": "John",
        ...     "first_name_corrected": "John",
        ...     "gender": "M",
        ...     "gender_corrected": "1"
        ... }
        >>> pairs = get_field_pairs(row)
        >>> pairs
        [('first_name', 'first_name_corrected'), ('gender', 'gender_corrected')]
    
    Requirements:
        - Validates: Requirements 13.2-13.4, 19.4
    """
    pairs = []
    original_fields = [key for key in json_row.keys() if not is_corrected_field(key)]
    
    for original_field in original_fields:
        corrected_field = get_corrected_field_name(original_field)
        if corrected_field in json_row:
            pairs.append((original_field, corrected_field))
    
    return pairs


def validate_field_naming_convention(json_row: JsonRow) -> tuple[bool, Optional[List[str]]]:
    """Validate that a JsonRow follows the field naming convention.
    
    Checks:
        - All field names use snake_case
        - Corrected fields have "_corrected" suffix
        - Every original field has a corresponding corrected field
        - No orphaned corrected fields (corrected field without original)
    
    Args:
        json_row: JsonRow dictionary to validate
    
    Returns:
        Tuple of (is_valid, error_messages)
        - is_valid: True if naming convention is followed, False otherwise
        - error_messages: List of validation error messages, None if valid
    
    Example:
        >>> row = {"first_name": "John", "first_name_corrected": "John"}
        >>> is_valid, errors = validate_field_naming_convention(row)
        >>> assert is_valid
    
    Requirements:
        - Validates: Requirements 13.3, 14.3-14.4, 19.4
    """
    errors = []
    
    # Check all field names use snake_case
    for field_name in json_row.keys():
        if not field_name.replace("_", "").isalnum():
            errors.append(f"Field name '{field_name}' does not follow snake_case convention")
    
    # Get original and corrected fields
    original_fields = [key for key in json_row.keys() if not is_corrected_field(key)]
    corrected_fields = [key for key in json_row.keys() if is_corrected_field(key)]
    
    # Check every original field has a corresponding corrected field
    for original_field in original_fields:
        corrected_field = get_corrected_field_name(original_field)
        if corrected_field not in json_row:
            errors.append(
                f"Original field '{original_field}' is missing corresponding "
                f"corrected field '{corrected_field}'"
            )
    
    # Check no orphaned corrected fields
    for corrected_field in corrected_fields:
        original_field = get_original_field_name(corrected_field)
        if original_field not in json_row:
            errors.append(
                f"Corrected field '{corrected_field}' has no corresponding "
                f"original field '{original_field}'"
            )
    
    if errors:
        return False, errors
    
    return True, None


# ============================================================================
# Validation Wrapper Functions
# ============================================================================

def validate_sheet_dataset_schema(
    dataset: SheetDataset,
    raise_on_error: bool = False
) -> tuple[bool, Optional[List[str]]]:
    """Validate a SheetDataset against its JSON schema.
    
    This function validates both the structure and the field naming convention.
    
    Args:
        dataset: SheetDataset instance to validate
        raise_on_error: If True, raise ValidationError on validation failure
    
    Returns:
        Tuple of (is_valid, error_messages)
        - is_valid: True if validation passed, False otherwise
        - error_messages: List of validation error messages, None if valid
    
    Raises:
        ValidationError: If raise_on_error is True and validation fails
        ImportError: If jsonschema library is not installed
    
    Example:
        >>> dataset = SheetDataset(
        ...     sheet_name="Students",
        ...     header_row=1,
        ...     header_rows_count=1,
        ...     field_names=["first_name"],
        ...     rows=[{"first_name": "John", "first_name_corrected": "John"}],
        ...     metadata={}
        ... )
        >>> is_valid, errors = validate_sheet_dataset_schema(dataset)
        >>> assert is_valid
    
    Requirements:
        - Validates: Requirements 19.1-19.5
    """
    if not JSONSCHEMA_AVAILABLE:
        raise ImportError(
            "jsonschema library is required for schema validation. "
            "Install it with: pip install jsonschema"
        )
    
    try:
        schema = load_schema("sheet_dataset.schema.json")
        
        # Convert dataclass to dictionary for validation
        dataset_dict = {
            "sheet_name": dataset.sheet_name,
            "header_row": dataset.header_row,
            "header_rows_count": dataset.header_rows_count,
            "field_names": dataset.field_names,
            "rows": dataset.rows,
            "metadata": dataset.metadata
        }
        
        validator = Draft7Validator(schema)
        
        # Collect all validation errors
        errors = list(validator.iter_errors(dataset_dict))
        
        if errors:
            if raise_on_error:
                raise ValidationError(f"SheetDataset validation failed: {errors[0].message}")
            
            error_messages = [
                f"{'.'.join(str(p) for p in error.path)}: {error.message}"
                for error in errors
            ]
            return False, error_messages
        
        # Also validate field naming convention for all rows
        for i, row in enumerate(dataset.rows):
            is_valid, naming_errors = validate_field_naming_convention(row)
            if not is_valid:
                if raise_on_error:
                    raise ValidationError(
                        f"Row {i} field naming validation failed: {naming_errors[0]}"
                    )
                return False, [f"Row {i}: {err}" for err in naming_errors]
        
        return True, None
    
    except (FileNotFoundError, json.JSONDecodeError) as e:
        if raise_on_error:
            raise
        return False, [str(e)]


def validate_workbook_dataset_schema(
    dataset: WorkbookDataset,
    raise_on_error: bool = False
) -> tuple[bool, Optional[List[str]]]:
    """Validate a WorkbookDataset against its JSON schema.
    
    This function validates both the structure and all contained SheetDatasets.
    
    Args:
        dataset: WorkbookDataset instance to validate
        raise_on_error: If True, raise ValidationError on validation failure
    
    Returns:
        Tuple of (is_valid, error_messages)
        - is_valid: True if validation passed, False otherwise
        - error_messages: List of validation error messages, None if valid
    
    Raises:
        ValidationError: If raise_on_error is True and validation fails
        ImportError: If jsonschema library is not installed
    
    Example:
        >>> dataset = WorkbookDataset(
        ...     source_file="data.xlsx",
        ...     sheets=[],
        ...     metadata={}
        ... )
        >>> is_valid, errors = validate_workbook_dataset_schema(dataset)
        >>> assert is_valid
    
    Requirements:
        - Validates: Requirements 19.1-19.5
    """
    if not JSONSCHEMA_AVAILABLE:
        raise ImportError(
            "jsonschema library is required for schema validation. "
            "Install it with: pip install jsonschema"
        )
    
    try:
        schema = load_schema("workbook_dataset.schema.json")
        
        # Convert dataclass to dictionary for validation
        sheets_dict = []
        for sheet in dataset.sheets:
            sheets_dict.append({
                "sheet_name": sheet.sheet_name,
                "header_row": sheet.header_row,
                "header_rows_count": sheet.header_rows_count,
                "field_names": sheet.field_names,
                "rows": sheet.rows,
                "metadata": sheet.metadata
            })
        
        dataset_dict = {
            "source_file": dataset.source_file,
            "sheets": sheets_dict,
            "metadata": dataset.metadata
        }
        
        validator = Draft7Validator(schema)
        
        # Collect all validation errors
        errors = list(validator.iter_errors(dataset_dict))
        
        if errors:
            if raise_on_error:
                raise ValidationError(f"WorkbookDataset validation failed: {errors[0].message}")
            
            error_messages = [
                f"{'.'.join(str(p) for p in error.path)}: {error.message}"
                for error in errors
            ]
            return False, error_messages
        
        # Also validate each sheet
        for i, sheet in enumerate(dataset.sheets):
            is_valid, sheet_errors = validate_sheet_dataset_schema(sheet, raise_on_error=False)
            if not is_valid:
                if raise_on_error:
                    raise ValidationError(
                        f"Sheet {i} ({sheet.sheet_name}) validation failed: {sheet_errors[0]}"
                    )
                return False, [f"Sheet {i} ({sheet.sheet_name}): {err}" for err in sheet_errors]
        
        return True, None
    
    except (FileNotFoundError, json.JSONDecodeError) as e:
        if raise_on_error:
            raise
        return False, [str(e)]


# ============================================================================
# Convenience Functions
# ============================================================================

def is_jsonschema_available() -> bool:
    """Check if jsonschema library is available.
    
    Returns:
        True if jsonschema is installed, False otherwise
    """
    return JSONSCHEMA_AVAILABLE


def get_available_schemas() -> List[str]:
    """Get list of available schema files.
    
    Returns:
        List of schema file names
    """
    schema_dir = _get_schema_path("").parent
    if not schema_dir.exists():
        return []
    
    return [f.name for f in schema_dir.glob("*.schema.json")]
