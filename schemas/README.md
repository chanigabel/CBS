# JSON Schema Documentation

This directory contains JSON Schema definitions for the Excel normalization data structures. These schemas define the structure, types, and validation rules for the JSON-based data pipeline.

## Overview

The Excel normalization system uses JSON as the internal data representation to decouple Excel complexity from normalization logic. The schemas define three core data structures:

1. **JsonRow**: A single data row with original and corrected field values
2. **SheetDataset**: All data from a single worksheet
3. **WorkbookDataset**: All data from an entire workbook

## Schema Files

### json_row.schema.json

Defines the structure of a single data row extracted from Excel.

**Key Features**:
- Supports both original and corrected field values
- Follows strict field naming convention
- Supports multiple data types (string, integer, null)
- Handles both single-column and split-column date fields

**Field Naming Convention**:
- Original fields: `field_name` (e.g., `first_name`, `gender`)
- Corrected fields: `field_name_corrected` (e.g., `first_name_corrected`, `gender_corrected`)

**Supported Fields**:

| Field Name | Type | Description |
|------------|------|-------------|
| `first_name` | string \| null | Original first name from Excel |
| `first_name_corrected` | string \| null | Normalized first name |
| `last_name` | string \| null | Original last name from Excel |
| `last_name_corrected` | string \| null | Normalized last name |
| `father_name` | string \| null | Original father name from Excel |
| `father_name_corrected` | string \| null | Normalized father name |
| `gender` | string \| integer \| null | Original gender (Hebrew/English text or code) |
| `gender_corrected` | string \| integer \| null | Normalized gender code |
| `id_number` | string \| null | Original ID number from Excel |
| `id_number_corrected` | string \| null | Normalized ID number |
| `passport` | string \| null | Original passport from Excel |
| `passport_corrected` | string \| null | Normalized passport |
| `birth_date` | string \| null | Original birth date (single column) |
| `birth_date_corrected` | string \| null | Normalized birth date |
| `birth_year` | integer \| null | Original birth year (split date) |
| `birth_year_corrected` | integer \| null | Normalized birth year |
| `birth_month` | integer \| null | Original birth month (split date) |
| `birth_month_corrected` | integer \| null | Normalized birth month |
| `birth_day` | integer \| null | Original birth day (split date) |
| `birth_day_corrected` | integer \| null | Normalized birth day |
| `entry_date` | string \| null | Original entry date (single column) |
| `entry_date_corrected` | string \| null | Normalized entry date |
| `entry_year` | integer \| null | Original entry year (split date) |
| `entry_year_corrected` | integer \| null | Normalized entry year |
| `entry_month` | integer \| null | Original entry month (split date) |
| `entry_month_corrected` | integer \| null | Normalized entry month |
| `entry_day` | integer \| null | Original entry day (split date) |
| `entry_day_corrected` | integer \| null | Normalized entry day |

**Example**:
```json
{
  "first_name": "יוסי",
  "first_name_corrected": "יוסי",
  "last_name": "כהן",
  "last_name_corrected": "כהן",
  "gender": "ז",
  "gender_corrected": "2",
  "birth_year": 1980,
  "birth_year_corrected": 1980,
  "birth_month": 5,
  "birth_month_corrected": 5,
  "birth_day": 15,
  "birth_day_corrected": 15
}
```

### sheet_dataset.schema.json

Defines the structure of a dataset for a single worksheet.

**Properties**:

| Property | Type | Required | Description |
|----------|------|----------|-------------|
| `sheet_name` | string | Yes | Name of the worksheet |
| `header_row` | integer | Yes | Row number where headers were found (1-based, minimum: 1) |
| `header_rows_count` | integer | Yes | Number of header rows (1 or 2) |
| `field_names` | array[string] | Yes | List of detected field names (unique, non-empty) |
| `rows` | array[JsonRow] | Yes | List of JSON row dictionaries |
| `metadata` | object | No | Additional metadata about the sheet |

**Metadata Properties**:

| Property | Type | Description |
|----------|------|-------------|
| `source_file` | string | Path to source Excel file |
| `extraction_date` | string | Date when data was extracted (ISO 8601) |
| `total_rows` | integer | Total number of data rows |
| `date_field_structure` | object | Mapping of date fields to "single" or "split" |
| `skipped_rows` | integer | Number of rows skipped during extraction |
| `errors` | array[string] | List of error messages |

**Example**:
```json
{
  "sheet_name": "Students",
  "header_row": 2,
  "header_rows_count": 1,
  "field_names": ["first_name", "last_name", "gender"],
  "rows": [
    {
      "first_name": "יוסי",
      "first_name_corrected": "יוסי",
      "last_name": "כהן",
      "last_name_corrected": "כהן",
      "gender": "ז",
      "gender_corrected": "2"
    }
  ],
  "metadata": {
    "source_file": "data.xlsx",
    "extraction_date": "2024-01-15",
    "total_rows": 1
  }
}
```

### workbook_dataset.schema.json

Defines the structure of a dataset for an entire workbook.

**Properties**:

| Property | Type | Required | Description |
|----------|------|----------|-------------|
| `source_file` | string | Yes | Path to the source Excel file |
| `sheets` | array[SheetDataset] | Yes | List of SheetDataset instances |
| `metadata` | object | No | Workbook-level metadata |

**Metadata Properties**:

| Property | Type | Description |
|----------|------|-------------|
| `extraction_date` | string | Date when data was extracted (ISO 8601) |
| `total_sheets` | integer | Total number of sheets in workbook |
| `processed_sheets` | integer | Number of sheets successfully processed |
| `skipped_sheets` | array[string] | List of sheet names that were skipped |
| `errors` | array[string] | List of error messages |

**Example**:
```json
{
  "source_file": "data.xlsx",
  "sheets": [
    {
      "sheet_name": "Students",
      "header_row": 1,
      "header_rows_count": 1,
      "field_names": ["first_name", "last_name"],
      "rows": [
        {
          "first_name": "יוסי",
          "first_name_corrected": "יוסי",
          "last_name": "כהן",
          "last_name_corrected": "כהן"
        }
      ],
      "metadata": {}
    }
  ],
  "metadata": {
    "extraction_date": "2024-01-15",
    "total_sheets": 1,
    "processed_sheets": 1,
    "skipped_sheets": []
  }
}
```

## Field Naming Convention

The system follows a strict naming convention for fields in JsonRow dictionaries:

### Original Fields

Original fields contain values extracted directly from Excel without modification:
- Use snake_case naming (lowercase with underscores)
- Examples: `first_name`, `last_name`, `gender`, `birth_year`

### Corrected Fields

Corrected fields contain normalized values after applying normalization engines:
- Use snake_case naming with `_corrected` suffix
- Examples: `first_name_corrected`, `last_name_corrected`, `gender_corrected`

### Convention Rules

1. **Every original field MUST have a corresponding corrected field**
   - If `first_name` exists, `first_name_corrected` must also exist
   - Even if no normalization was applied, the corrected field must be present

2. **No orphaned corrected fields**
   - A corrected field cannot exist without its original field
   - `first_name_corrected` requires `first_name` to exist

3. **Original values are never modified**
   - Original fields preserve exact values from Excel
   - Normalization results are stored only in corrected fields

4. **Corrected fields may contain original values**
   - If normalization fails or is not applicable, corrected field contains original value
   - This ensures corrected fields always have a value when original field has a value

### Date Field Naming

Date fields can be represented in two ways:

**Single Column Dates**:
- `birth_date` / `birth_date_corrected`
- `entry_date` / `entry_date_corrected`

**Split Column Dates**:
- `birth_year` / `birth_year_corrected`
- `birth_month` / `birth_month_corrected`
- `birth_day` / `birth_day_corrected`
- `entry_year` / `entry_year_corrected`
- `entry_month` / `entry_month_corrected`
- `entry_day` / `entry_day_corrected`

A JsonRow will contain either single or split date fields for each date type, but not both.

## Using the Validation Utilities

The `schema_validation.py` module provides utilities to validate data structures against these schemas.

### Installation

Schema validation requires the `jsonschema` library:

```bash
pip install jsonschema
```

### Basic Usage

```python
from excel_normalization.schema_validation import (
    validate_json_row,
    validate_sheet_dataset_schema,
    validate_workbook_dataset_schema,
    validate_field_naming_convention
)
from excel_normalization.data_types import SheetDataset, JsonRow

# Validate a JsonRow
row = {"first_name": "John", "first_name_corrected": "John"}
is_valid, errors = validate_json_row(row)
if not is_valid:
    print(f"Validation errors: {errors}")

# Validate a SheetDataset
dataset = SheetDataset(
    sheet_name="Students",
    header_row=1,
    header_rows_count=1,
    field_names=["first_name"],
    rows=[row],
    metadata={}
)
is_valid, errors = validate_sheet_dataset_schema(dataset)
if not is_valid:
    print(f"Validation errors: {errors}")

# Validate field naming convention
is_valid, errors = validate_field_naming_convention(row)
if not is_valid:
    print(f"Naming convention errors: {errors}")
```

### Validation Options

All validation functions support two modes:

1. **Return errors** (default): Returns `(is_valid, error_messages)` tuple
2. **Raise on error**: Set `raise_on_error=True` to raise `ValidationError` on failure

```python
# Return errors mode
is_valid, errors = validate_sheet_dataset_schema(dataset)

# Raise on error mode
try:
    validate_sheet_dataset_schema(dataset, raise_on_error=True)
except ValidationError as e:
    print(f"Validation failed: {e}")
```

### Field Naming Utilities

The module provides utilities for working with the field naming convention:

```python
from excel_normalization.schema_validation import (
    get_corrected_field_name,
    get_original_field_name,
    is_corrected_field,
    get_field_pairs
)

# Get corrected field name
corrected = get_corrected_field_name("first_name")  # "first_name_corrected"

# Get original field name
original = get_original_field_name("first_name_corrected")  # "first_name"

# Check if field is corrected
is_corrected = is_corrected_field("first_name_corrected")  # True

# Get all field pairs from a row
row = {
    "first_name": "John",
    "first_name_corrected": "John",
    "gender": "M",
    "gender_corrected": "1"
}
pairs = get_field_pairs(row)
# [('first_name', 'first_name_corrected'), ('gender', 'gender_corrected')]
```

## Schema Validation Rules

### JsonRow Validation

- All field names must use snake_case convention
- Corrected fields must have `_corrected` suffix
- Field values must match expected types (string, integer, or null)
- No additional properties beyond defined fields

### SheetDataset Validation

- `sheet_name` must be non-empty string
- `header_row` must be positive integer (≥1)
- `header_rows_count` must be 1 or 2
- `field_names` must be non-empty array with unique values
- `rows` must be array of valid JsonRow objects
- All rows must follow field naming convention

### WorkbookDataset Validation

- `source_file` must be non-empty string
- `sheets` must be array of valid SheetDataset objects
- All sheet names must be unique
- All sheets must pass SheetDataset validation

## Requirements Traceability

These schemas and validation utilities satisfy the following requirements:

- **Requirement 19.1**: JSON schema for Sheet_Dataset defined
- **Requirement 19.2**: Required and optional fields specified
- **Requirement 19.3**: Field types defined (string, number, date, etc.)
- **Requirement 19.4**: Corrected field naming convention documented
- **Requirement 19.5**: JSON schema available in documentation

Additional requirements validated:
- **Requirements 10.2, 14.3-14.4**: JsonRow structure and field naming
- **Requirements 11.2-11.5**: SheetDataset structure and metadata
- **Requirements 13.2-13.5**: Corrected field creation and naming
- **Requirements 16.1-16.4**: WorkbookDataset structure for multi-sheet support

## Integration with Python Code

The schemas are designed to match the Python dataclasses defined in `src/excel_normalization/data_types.py`:

- `JsonRow` → Python type alias `Dict[str, Any]`
- `SheetDataset` → Python `@dataclass SheetDataset`
- `WorkbookDataset` → Python `@dataclass WorkbookDataset`

The validation utilities in `src/excel_normalization/schema_validation.py` bridge the gap between Python dataclasses and JSON schemas, allowing runtime validation of data structures.

## Usage in Testing

These schemas are particularly useful for:

1. **Property-based testing**: Generate random valid data structures
2. **Integration testing**: Validate pipeline outputs
3. **API testing**: Validate JSON exports and imports
4. **Documentation**: Provide clear contract for data structures

Example property-based test:

```python
from hypothesis import given
from hypothesis_jsonschema import from_schema
from excel_normalization.schema_validation import load_schema, validate_json_row

# Generate random JsonRow instances that conform to schema
json_row_schema = load_schema("json_row.schema.json")

@given(from_schema(json_row_schema))
def test_json_row_always_valid(row):
    """Generated rows should always pass validation."""
    is_valid, errors = validate_json_row(row)
    assert is_valid, f"Generated row failed validation: {errors}"
```

## Schema Versioning

Current schema version: **1.0.0**

The schemas follow semantic versioning:
- **Major version**: Breaking changes to schema structure
- **Minor version**: Backward-compatible additions
- **Patch version**: Documentation or clarification updates

## Future Enhancements

Potential future additions to the schemas:

1. **Validation rules**: Add pattern matching for ID numbers, date formats
2. **Custom fields**: Support for institution-specific custom fields
3. **Localization**: Support for additional languages beyond Hebrew/English
4. **Performance metadata**: Add timing and performance metrics to metadata
5. **Data quality metrics**: Add validation scores and confidence levels

## References

- JSON Schema Specification: https://json-schema.org/
- Draft 7 Schema: https://json-schema.org/draft-07/schema
- Python jsonschema library: https://python-jsonschema.readthedocs.io/
