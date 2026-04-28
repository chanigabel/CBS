# Field Naming Convention

This document describes the field naming convention used throughout the Excel standardization system for JSON data structures.

## Overview

The system uses a consistent naming pattern to distinguish between original values extracted from Excel and corrected values produced by standardization engines.

## Convention Rules

### 1. Original Fields

Original fields contain values extracted directly from Excel without any modification.

**Format**: `field_name` (snake_case)

**Examples**:
- `first_name`
- `last_name`
- `gender`
- `birth_year`
- `id_number`

**Characteristics**:
- Use lowercase letters and underscores only
- No special characters or spaces
- Descriptive and standardized across all institutions
- Preserve exact values from Excel cells

### 2. Corrected Fields

Corrected fields contain normalized values after applying standardization engines.

**Format**: `field_name_corrected` (snake_case with `_corrected` suffix)

**Examples**:
- `first_name_corrected`
- `last_name_corrected`
- `gender_corrected`
- `birth_year_corrected`
- `id_number_corrected`

**Characteristics**:
- Always paired with an original field
- Use same base name as original field plus `_corrected` suffix
- May contain normalized value or original value if standardization failed
- Never exist without corresponding original field

### 3. Pairing Requirement

**Every original field MUST have a corresponding corrected field.**

This ensures:
- Consistent data structure across all rows
- Easy comparison between original and corrected values
- Clear indication of which fields have been processed
- Predictable field names for downstream processing

**Valid Example**:
```json
{
  "first_name": "יוסי",
  "first_name_corrected": "יוסי",
  "gender": "ז",
  "gender_corrected": "2"
}
```

**Invalid Example** (missing corrected field):
```json
{
  "first_name": "יוסי",
  "gender": "ז",
  "gender_corrected": "2"
}
```

### 4. No Orphaned Corrected Fields

**A corrected field cannot exist without its original field.**

This prevents:
- Confusion about data source
- Loss of original values
- Incomplete data structures

**Invalid Example** (orphaned corrected field):
```json
{
  "first_name": "יוסי",
  "first_name_corrected": "יוסי",
  "gender_corrected": "2"
}
```

## Supported Field Names

### Name Fields

| Original Field | Corrected Field | Type | Description |
|----------------|-----------------|------|-------------|
| `first_name` | `first_name_corrected` | string \| null | First name |
| `last_name` | `last_name_corrected` | string \| null | Last name |
| `father_name` | `father_name_corrected` | string \| null | Father's name |

### Gender Field

| Original Field | Corrected Field | Type | Description |
|----------------|-----------------|------|-------------|
| `gender` | `gender_corrected` | string \| integer \| null | Gender (text or code) |

### Identifier Fields

| Original Field | Corrected Field | Type | Description |
|----------------|-----------------|------|-------------|
| `id_number` | `id_number_corrected` | string \| null | Israeli ID number |
| `passport` | `passport_corrected` | string \| null | Passport number |

### Date Fields (Single Column)

| Original Field | Corrected Field | Type | Description |
|----------------|-----------------|------|-------------|
| `birth_date` | `birth_date_corrected` | string \| null | Birth date |
| `entry_date` | `entry_date_corrected` | string \| null | Entry date |

### Date Fields (Split Columns)

| Original Field | Corrected Field | Type | Description |
|----------------|-----------------|------|-------------|
| `birth_year` | `birth_year_corrected` | integer \| null | Birth year |
| `birth_month` | `birth_month_corrected` | integer \| null | Birth month |
| `birth_day` | `birth_day_corrected` | integer \| null | Birth day |
| `entry_year` | `entry_year_corrected` | integer \| null | Entry year |
| `entry_month` | `entry_month_corrected` | integer \| null | Entry month |
| `entry_day` | `entry_day_corrected` | integer \| null | Entry day |

## Date Field Structure

Date fields can be represented in two mutually exclusive ways:

### Single Column Dates

When a date appears in one Excel column:
- Use `birth_date` / `birth_date_corrected`
- Use `entry_date` / `entry_date_corrected`
- Values are typically strings in various formats

### Split Column Dates

When a date is split across multiple Excel columns:
- Use `birth_year`, `birth_month`, `birth_day` with their corrected counterparts
- Use `entry_year`, `entry_month`, `entry_day` with their corrected counterparts
- Values are integers

**Important**: A JsonRow will contain either single OR split date fields for each date type, never both.

## Implementation Guidelines

### When Creating JsonRow

```python
# Always create both original and corrected fields
json_row = {
    "first_name": original_value,
    "first_name_corrected": None  # Will be filled by standardization
}
```

### When standardizing

```python
# Apply standardization and update corrected field
if "first_name" in json_row:
    original = json_row["first_name"]
    corrected = name_engine.normalize_name(original)
    json_row["first_name_corrected"] = corrected
```

### When Handling Errors

```python
# If standardization fails, use original value
try:
    corrected = engine.normalize(original)
    json_row["field_name_corrected"] = corrected
except Exception:
    json_row["field_name_corrected"] = original
```

## Validation

Use the validation utilities to ensure compliance:

```python
from src.excel_standardization.schema_validation import (
    validate_field_naming_convention,
    get_corrected_field_name,
    is_corrected_field
)

# Validate a row follows the convention
is_valid, errors = validate_field_naming_convention(json_row)
if not is_valid:
    print(f"Naming convention errors: {errors}")

# Get corrected field name
corrected_name = get_corrected_field_name("first_name")  # "first_name_corrected"

# Check if field is corrected
if is_corrected_field("first_name_corrected"):
    print("This is a corrected field")
```

## Benefits of This Convention

1. **Clarity**: Immediately clear which fields are original vs corrected
2. **Consistency**: Same pattern across all field types
3. **Traceability**: Can always trace back to original values
4. **Validation**: Easy to validate programmatically
5. **Export**: Simple to create Excel columns with both original and corrected data
6. **Non-destructive**: Original data is never lost or overwritten

## Requirements Traceability

This naming convention satisfies:
- **Requirement 13.3**: Corrected field naming with "_corrected" suffix
- **Requirement 14.3**: Original values remain in "field_name" key
- **Requirement 14.4**: Normalized values stored in "field_name_corrected" key
- **Requirement 19.4**: Corrected field naming convention documented
