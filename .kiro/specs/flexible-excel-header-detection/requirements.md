# Requirements Document: JSON-Based Excel Data Normalization Pipeline

## Introduction

This document specifies requirements for a JSON-based data normalization pipeline that processes Excel worksheets from multiple institutions with varying structures. The system detects column headers, extracts data rows, converts them to JSON format, applies normalization engines, and produces both raw and corrected JSON datasets that can be exported back to Excel.

## Glossary

- **Header_Detector**: The component responsible for scanning Excel worksheets and identifying header row locations and column mappings
- **ExcelReader**: The existing layer that reads Excel files using openpyxl library
- **ExcelToJsonExtractor**: Component that converts Excel rows to JSON format based on column mappings
- **RawJsonDataset**: JSON representation of original Excel data without modifications
- **CorrectedJsonDataset**: JSON representation with normalized values from engines
- **NormalizationPipeline**: Component that applies normalization engines to JSON rows
- **JsonToExcelWriter**: Component that exports corrected JSON back to Excel format
- **Normalized_Text**: Text that has been processed to remove line breaks, parentheses, extra whitespace, and converted to lowercase
- **Keyword_Pattern**: A set of Hebrew or English terms that identify a specific field type (e.g., "שם פרטי", "first name")
- **Header_Row**: The row(s) in an Excel worksheet containing column names that describe the data below
- **Column_Mapping**: A dictionary mapping standardized field names to Excel column indices
- **Split_Date_Field**: A date represented across multiple columns (year, month, day)
- **Single_Date_Field**: A date represented in one column
- **Corrected_Column**: A column containing the Hebrew word "מתוקן" indicating corrected data that should be ignored
- **Merged_Cell**: An Excel cell that spans multiple rows or columns
- **Scan_Range**: The first 30 rows of a worksheet where headers are expected to appear
- **JSON_Row**: A dictionary representing a single data row with field names as keys and values from Excel cells
- **Corrected_Field**: A field in JSON containing normalized value (e.g., "gender_corrected" for normalized gender)
- **Sheet_Dataset**: A JSON structure containing all rows from a single worksheet with both raw and corrected values

## Requirements

### Requirement 1: Scan Worksheet for Header Row

**User Story:** As a data processor, I want the system to scan the beginning of each worksheet, so that headers can be found regardless of their position

#### Acceptance Criteria

1. THE Header_Detector SHALL scan the first 30 rows of each worksheet
2. WHEN a worksheet has fewer than 30 rows, THE Header_Detector SHALL scan all available rows
3. FOR ALL rows in the Scan_Range, THE Header_Detector SHALL extract cell values for analysis
4. THE Header_Detector SHALL handle Merged_Cells without raising exceptions
5. THE Header_Detector SHALL identify whether headers span one or two rows

### Requirement 2: Normalize Cell Text

**User Story:** As a data processor, I want cell text to be normalized before matching, so that variations in formatting do not prevent header detection

#### Acceptance Criteria

1. WHEN processing cell text, THE Header_Detector SHALL remove all line breaks from the text
2. WHEN processing cell text, THE Header_Detector SHALL remove all parentheses from the text
3. WHEN processing cell text, THE Header_Detector SHALL collapse multiple consecutive whitespace characters into single spaces
4. WHEN processing cell text, THE Header_Detector SHALL convert all text to lowercase
5. WHEN processing cell text, THE Header_Detector SHALL trim leading and trailing whitespace
6. THE Header_Detector SHALL produce Normalized_Text for all cell values before pattern matching

### Requirement 3: Support Multilingual Keyword Patterns

**User Story:** As a data processor, I want the system to recognize both Hebrew and English column names, so that files from different institutions can be processed

#### Acceptance Criteria

1. THE Header_Detector SHALL define Keyword_Patterns for first name in both Hebrew and English
2. THE Header_Detector SHALL define Keyword_Patterns for last name in both Hebrew and English
3. THE Header_Detector SHALL define Keyword_Patterns for father name in both Hebrew and English
4. THE Header_Detector SHALL define Keyword_Patterns for gender in both Hebrew and English
5. THE Header_Detector SHALL define Keyword_Patterns for ID number in both Hebrew and English
6. THE Header_Detector SHALL define Keyword_Patterns for passport in both Hebrew and English
7. THE Header_Detector SHALL define Keyword_Patterns for birth date fields in both Hebrew and English
8. THE Header_Detector SHALL define Keyword_Patterns for entry date in both Hebrew and English
9. WHEN matching cell text, THE Header_Detector SHALL use substring matching rather than exact equality
10. THE Header_Detector SHALL match Normalized_Text against normalized Keyword_Patterns

### Requirement 4: Ignore Corrected Columns

**User Story:** As a data processor, I want columns marked as corrected to be ignored, so that only original data columns are mapped

#### Acceptance Criteria

1. WHEN a cell contains the Hebrew word "מתוקן", THE Header_Detector SHALL mark that column as a Corrected_Column
2. THE Header_Detector SHALL exclude Corrected_Columns from the Column_Mapping
3. WHEN multiple columns match the same field pattern, THE Header_Detector SHALL prefer the column that is not a Corrected_Column

### Requirement 5: Identify Most Likely Header Row

**User Story:** As a data processor, I want the system to identify which row is the header, so that column mappings are accurate

#### Acceptance Criteria

1. FOR ALL rows in the Scan_Range, THE Header_Detector SHALL count how many cells match recognized Keyword_Patterns
2. THE Header_Detector SHALL select the row with the highest count of pattern matches as the Header_Row
3. WHEN multiple rows have equal match counts, THE Header_Detector SHALL select the first row among them
4. WHEN no row contains any pattern matches, THE Header_Detector SHALL return an empty Column_Mapping
5. THE Header_Detector SHALL require at least 3 pattern matches for a row to be considered a valid Header_Row

### Requirement 6: Support Single-Column Date Fields

**User Story:** As a data processor, I want to detect date fields in a single column, so that dates can be extracted from consolidated date columns

#### Acceptance Criteria

1. WHEN a column matches birth date Keyword_Patterns, THE Header_Detector SHALL map it to "birth_date" in the Column_Mapping
2. WHEN a column matches entry date Keyword_Patterns, THE Header_Detector SHALL map it to "entry_date" in the Column_Mapping
3. THE Header_Detector SHALL recognize Single_Date_Field patterns for both birth date and entry date

### Requirement 7: Support Split Date Fields

**User Story:** As a data processor, I want to detect date fields split across multiple columns, so that year, month, and day can be extracted separately

#### Acceptance Criteria

1. WHEN a column matches year Keyword_Patterns, THE Header_Detector SHALL map it to "birth_year" or "entry_year" based on context
2. WHEN a column matches month Keyword_Patterns, THE Header_Detector SHALL map it to "birth_month" or "entry_month" based on context
3. WHEN a column matches day Keyword_Patterns, THE Header_Detector SHALL map it to "birth_day" or "entry_day" based on context
4. THE Header_Detector SHALL recognize Split_Date_Field patterns in both Hebrew and English
5. WHEN both Single_Date_Field and Split_Date_Field patterns are detected for the same date type, THE Header_Detector SHALL prefer the Split_Date_Field mapping

### Requirement 8: Generate Column Mapping

**User Story:** As a data processor, I want a standardized mapping of field names to column indices, so that data can be extracted consistently

#### Acceptance Criteria

1. THE Header_Detector SHALL return a Column_Mapping as a Python dictionary
2. THE Column_Mapping SHALL use standardized field names as keys: "first_name", "last_name", "father_name", "gender", "id_number", "passport"
3. THE Column_Mapping SHALL use zero-based column indices as values
4. WHEN a field is not detected, THE Header_Detector SHALL omit that field from the Column_Mapping
5. THE Column_Mapping SHALL include either Single_Date_Field keys or Split_Date_Field keys for each date type, but not both
6. THE Header_Detector SHALL return the Column_Mapping after identifying the Header_Row

### Requirement 9: Handle Multi-Row Headers

**User Story:** As a data processor, I want the system to handle headers that span multiple rows, so that grouped column structures can be processed

#### Acceptance Criteria

1. WHEN analyzing potential Header_Rows, THE Header_Detector SHALL consider cells from adjacent rows for context
2. WHEN a parent header cell spans multiple child columns, THE Header_Detector SHALL associate child column names with the parent context
3. THE Header_Detector SHALL detect Split_Date_Fields even when year, month, and day appear as child columns under a parent date header
4. THE Header_Detector SHALL handle Merged_Cells that span multiple rows without data loss

### Requirement 10: Extract Data Rows to JSON

**User Story:** As a data processor, I want all data rows extracted to JSON format, so that normalization engines can process structured data

#### Acceptance Criteria

1. THE ExcelToJsonExtractor SHALL extract all rows below the header row(s)
2. FOR EACH data row, THE ExcelToJsonExtractor SHALL create a JSON_Row dictionary
3. THE JSON_Row SHALL contain field names as keys based on the Column_Mapping
4. THE JSON_Row SHALL contain original cell values as values
5. THE ExcelToJsonExtractor SHALL preserve the exact original values without modification
6. WHEN a cell is empty, THE ExcelToJsonExtractor SHALL store None or empty string as the value
7. THE ExcelToJsonExtractor SHALL handle cells with formulas by extracting the calculated value

### Requirement 11: Produce Raw JSON Dataset

**User Story:** As a data processor, I want a raw JSON dataset preserving original values, so that I can trace back to source data

#### Acceptance Criteria

1. THE system SHALL produce a RawJsonDataset containing all extracted rows
2. THE RawJsonDataset SHALL be a list of JSON_Row dictionaries
3. THE RawJsonDataset SHALL preserve original values exactly as they appear in Excel
4. THE RawJsonDataset SHALL include metadata about the source worksheet (name, header row number)
5. THE RawJsonDataset SHALL be stored in memory for processing
6. THE system SHALL support exporting RawJsonDataset to a JSON file

### Requirement 12: Apply Normalization Engines to JSON Rows

**User Story:** As a data processor, I want normalization engines to operate on JSON rows, so that data can be cleaned without modifying Excel directly

#### Acceptance Criteria

1. THE NormalizationPipeline SHALL accept a RawJsonDataset as input
2. THE NormalizationPipeline SHALL apply existing normalization engines to each JSON_Row
3. THE NormalizationPipeline SHALL invoke NameEngine for first_name, last_name, and father_name fields
4. THE NormalizationPipeline SHALL invoke GenderEngine for gender field
5. THE NormalizationPipeline SHALL invoke DateEngine for date fields (birth_date, entry_date, or split date components)
6. THE NormalizationPipeline SHALL invoke IdentifierEngine for id_number and passport fields
7. THE NormalizationPipeline SHALL pass field values from JSON_Row to engines
8. THE engines SHALL remain unchanged and operate on string values from JSON

### Requirement 13: Produce Corrected JSON Dataset

**User Story:** As a data processor, I want a corrected JSON dataset with normalized values, so that I can see both original and corrected data

#### Acceptance Criteria

1. THE system SHALL produce a CorrectedJsonDataset containing normalized values
2. FOR EACH field in a JSON_Row, THE system SHALL create a corresponding corrected field
3. THE corrected field SHALL be named with "_corrected" suffix (e.g., "gender_corrected")
4. THE CorrectedJsonDataset SHALL contain both original and corrected values in the same row
5. WHEN a field is not normalized, THE corrected field SHALL contain the original value
6. THE CorrectedJsonDataset SHALL maintain the same row order as RawJsonDataset
7. THE CorrectedJsonDataset SHALL include the same metadata as RawJsonDataset

### Requirement 14: Preserve Original Values

**User Story:** As a data processor, I want original values never to be overwritten, so that I can always reference source data

#### Acceptance Criteria

1. THE system SHALL NEVER modify original field values in JSON rows
2. THE system SHALL store normalized values in separate corrected fields
3. FOR ANY field "field_name", THE original value SHALL remain in "field_name" key
4. FOR ANY field "field_name", THE normalized value SHALL be stored in "field_name_corrected" key
5. THE system SHALL maintain both original and corrected values throughout the pipeline

### Requirement 15: Export Corrected JSON to Excel

**User Story:** As a data processor, I want to export corrected JSON back to Excel, so that I can review and share normalized data

#### Acceptance Criteria

1. THE JsonToExcelWriter SHALL accept a CorrectedJsonDataset as input
2. THE JsonToExcelWriter SHALL create a new Excel workbook
3. THE JsonToExcelWriter SHALL create one worksheet per Sheet_Dataset
4. THE JsonToExcelWriter SHALL write header row with field names
5. THE JsonToExcelWriter SHALL write both original and corrected columns
6. THE JsonToExcelWriter SHALL use column naming convention: "field_name" and "field_name_corrected"
7. THE JsonToExcelWriter SHALL preserve row order from the JSON dataset
8. THE JsonToExcelWriter SHALL save the workbook to a specified file path

### Requirement 16: Support Multiple Worksheets

**User Story:** As a data processor, I want to process multiple worksheets in a workbook, so that I can handle complex Excel files

#### Acceptance Criteria

1. THE system SHALL process each worksheet in the workbook independently
2. FOR EACH worksheet, THE system SHALL detect headers separately
3. FOR EACH worksheet, THE system SHALL produce a separate Sheet_Dataset
4. THE system SHALL maintain worksheet names in the output
5. THE system SHALL handle worksheets with different header structures
6. THE system SHALL skip worksheets with no valid headers

### Requirement 17: Maintain Backward Compatibility

**User Story:** As a developer, I want existing engines to work without modification, so that the system remains maintainable

#### Acceptance Criteria

1. THE NameEngine SHALL continue to accept string values and return normalized strings
2. THE GenderEngine SHALL continue to accept string values and return normalized strings
3. THE DateEngine SHALL continue to accept date components and return normalized dates
4. THE IdentifierEngine SHALL continue to accept string values and return normalized strings
5. THE system SHALL adapt JSON values to engine input formats
6. THE system SHALL adapt engine outputs back to JSON format
7. THE engines SHALL NOT require any code changes

### Requirement 18: Handle Missing Fields

**User Story:** As a data processor, I want the system to handle missing fields gracefully, so that incomplete data doesn't cause failures

#### Acceptance Criteria

1. WHEN a field is not present in the Column_Mapping, THE system SHALL skip normalization for that field
2. WHEN a field value is None or empty, THE system SHALL pass it to the engine or skip normalization based on engine requirements
3. THE system SHALL NOT raise exceptions for missing optional fields
4. THE system SHALL include all detected fields in the JSON output, even if some rows have missing values

### Requirement 19: Provide JSON Schema

**User Story:** As a developer, I want a clear JSON schema, so that I can understand the data structure

#### Acceptance Criteria

1. THE system SHALL define a JSON schema for Sheet_Dataset
2. THE JSON schema SHALL specify required and optional fields
3. THE JSON schema SHALL define field types (string, number, date, etc.)
4. THE JSON schema SHALL document the corrected field naming convention
5. THE JSON schema SHALL be available in documentation

### Requirement 20: Maintain Clean IO Separation

**User Story:** As a developer, I want clean separation between IO and business logic, so that the system is maintainable

#### Acceptance Criteria

1. THE ExcelReader layer SHALL handle all Excel read operations
2. THE ExcelToJsonExtractor SHALL handle conversion from Excel to JSON
3. THE NormalizationPipeline SHALL handle business logic without Excel dependencies
4. THE JsonToExcelWriter SHALL handle all Excel write operations
5. THE engines SHALL operate only on JSON/dictionary data structures
6. THE system SHALL NOT mix Excel operations with normalization logic

