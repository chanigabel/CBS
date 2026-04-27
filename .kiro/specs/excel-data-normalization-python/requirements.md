# Requirements Document

## Introduction

This document specifies requirements for a Python-based Excel data normalization system that transforms person records (residents, staff, etc.) from various sources into a standardized format. The system replaces a legacy VBA implementation with strict behavioral equivalence.

The system processes Excel workbooks with inconsistent structures, mixed languages (Hebrew/English), varying date formats, and invalid identifiers. The system writes corrected values to new columns next to the original columns in the same workbook, highlighting changed cells in pink.

**Critical Design Principle**: The Python system SHALL replicate the exact behavior of the legacy VBA implementation. All normalization, validation, and transformation logic SHALL follow the same rules observed in the VBA code. The system is fully rule-based and deterministic with no machine learning components.

## Glossary

- **System**: The Excel Data Normalization Python application
- **Legacy_VBA_System**: The original VBA implementation that this Python system replaces
- **Workbook**: An Excel file containing person records to be processed
- **Worksheet**: A single sheet within a Workbook
- **Header_Row**: The row in a Worksheet containing column names
- **Field_Processor**: A component that normalizes a specific field type (name, date, gender, identifier)
- **Identifier**: A unique person identifier (Israeli ID or passport number)
- **Israeli_ID**: A 9-digit identifier with checksum validation
- **Passport**: An alphanumeric travel document identifier
- **Orchestrator**: The component coordinating processing across Worksheets
- **Engine**: A specialized processor for a specific data type (Name_Engine, Date_Engine, etc.)
- **Language_Dominance**: The primary language (Hebrew/English) detected in text fields
- **Date_Format_Pattern**: A pattern for parsing date strings (DDMM or MMDD for ambiguous dates)
- **Checksum**: A validation digit calculated from other digits in Israeli_ID
- **Diacritic**: Accent marks or special characters removed during normalization
- **Hebrew_Final_Letter**: Hebrew letters that appear only at word endings (ך, ם, ן, ף, ץ)
- **Father_Name_Pattern**: A detected pattern for removing last name from father name field (RemoveFirst, RemoveLast, or None)
- **Corrected_Column**: A new column inserted next to the original column containing normalized values
- **Pink_Highlight**: Background color (RGB 255, 199, 206) applied to cells where the corrected value differs from the original


## Requirements

### Requirement 0: Legacy VBA Compatibility

**User Story:** As a system maintainer, I want the Python system to replicate VBA behavior exactly, so that I can ensure consistent results during the migration.

#### Acceptance Criteria

1. THE System SHALL replicate the behavior of the Legacy_VBA_System for all normalization operations
2. ALL normalization, validation, and transformation logic SHALL follow the same rules observed in the VBA code
3. THE System SHALL be fully rule-based and deterministic
4. THE System SHALL NOT use machine learning models or probabilistic algorithms
5. THE System SHALL produce functionally equivalent output to the Legacy_VBA_System for the same input data

### Requirement 1: Workbook Processing

**User Story:** As a data processor, I want to process all worksheets in a workbook, so that I can normalize data across multiple sheets.

#### Acceptance Criteria

1. WHEN a Workbook path is provided, THE System SHALL load the workbook
2. THE System SHALL process each Worksheet in the Workbook sequentially
3. FOR EACH Worksheet, THE System SHALL process names, gender, dates, and identifiers in that order
4. THE System SHALL modify the Workbook in place by inserting Corrected_Columns
5. THE System SHALL save the modified Workbook
6. THE System SHALL preserve all original data in the original columns

### Requirement 2: Header Detection

**User Story:** As a data processor, I want the system to find columns by exact header text matching, so that I can process worksheets with known column names.

#### Acceptance Criteria

1. WHEN searching for a column, THE System SHALL use exact text matching with partial match (xlPart)
2. THE System SHALL support multiple header variants for each field (e.g., "שם פרטי" or "first name")
3. WHEN a header is found, THE System SHALL record the column number, header row number, and last data row
4. WHEN a header is not found, THE System SHALL skip processing for that field
5. THE System SHALL NOT infer column types from data patterns
6. THE System SHALL NOT use confidence scores for header matching
7. THE header matching logic SHALL replicate the FindHeader behavior from the Legacy_VBA_System

### Requirement 3: Text Normalization

**User Story:** As a data processor, I want to normalize text fields consistently, so that I can produce clean standardized output.

#### Acceptance Criteria

1. WHEN a text field is processed, THE System SHALL trim leading and trailing whitespace
2. THE System SHALL count Hebrew letters (Unicode 1488-1514) and English letters (A-Z, a-z)
3. THE System SHALL determine Language_Dominance by comparing Hebrew count to English count
4. THE System SHALL remove Diacritics using character code mappings
5. THE System SHALL keep only letters of the dominant language
6. WHEN processing Hebrew text, THE System SHALL add a space after Hebrew_Final_Letters if the next character is not a space or punctuation
7. THE System SHALL collapse multiple consecutive spaces into a single space
8. THE System SHALL treat spaces, hyphens, en-dashes, and em-dashes as valid separators
9. THE text normalization logic SHALL replicate the CleanName behavior from the Legacy_VBA_System

### Requirement 4: Name Field Processing

**User Story:** As a data processor, I want to normalize first name and last name fields, so that I can produce clean standardized output.

#### Acceptance Criteria

1. WHEN processing names, THE System SHALL search for "שם פרטי" or "first name" headers
2. WHEN processing names, THE System SHALL search for "שם משפחה" or "last name" headers
3. FOR EACH found name column, THE System SHALL insert a Corrected_Column immediately after the original column
4. THE Corrected_Column header SHALL be the original header text with " - מתוקן" appended
5. THE System SHALL apply text normalization to each name value
6. THE System SHALL write the normalized values to the Corrected_Column
7. WHEN a normalized value differs from the original value, THE System SHALL apply Pink_Highlight to the corrected cell
8. THE name processing logic SHALL replicate the ProcessNames behavior from the Legacy_VBA_System

### Requirement 5: Father Name Pattern Detection

**User Story:** As a data processor, I want to detect patterns in father names that include the last name, so that I can remove the last name from the father name field.

#### Acceptance Criteria

1. WHEN processing father names, THE System SHALL sample the first 5 rows
2. FOR EACH sampled row, THE System SHALL check if the father name contains the last name (case-sensitive binary comparison)
3. THE System SHALL count how many sampled father names contain the last name
4. WHEN fewer than 3 sampled father names contain the last name, THE System SHALL use Father_Name_Pattern None
5. WHEN 3 or more sampled father names contain the last name, THE System SHALL split each father name by spaces and check the position of the last name
6. WHEN 3 or more have the last name in the first position, THE System SHALL use Father_Name_Pattern RemoveFirst
7. WHEN 3 or more have the last name in the last position, THE System SHALL use Father_Name_Pattern RemoveLast
8. OTHERWISE, THE System SHALL use Father_Name_Pattern None
9. THE pattern detection logic SHALL replicate the DetectFatherNamePattern behavior from the Legacy_VBA_System

### Requirement 6: Father Name Processing

**User Story:** As a data processor, I want to normalize father names and remove the last name when appropriate, so that I can produce clean father name values.

#### Acceptance Criteria

1. WHEN processing father names, THE System SHALL search for "שם האב" or "father's name" headers
2. THE System SHALL insert a Corrected_Column immediately after the original column
3. THE System SHALL apply text normalization to each father name value
4. WHEN a last name column is found, THE System SHALL detect the Father_Name_Pattern
5. WHEN the Father_Name_Pattern is not None, THE System SHALL remove the last name substring from the father name
6. WHEN the Father_Name_Pattern is RemoveFirst, THE System SHALL remove the first word after removing the last name substring
7. WHEN the Father_Name_Pattern is RemoveLast, THE System SHALL remove the last word after removing the last name substring
8. THE System SHALL write the normalized values to the Corrected_Column
9. WHEN a normalized value differs from the original value, THE System SHALL apply Pink_Highlight to the corrected cell
10. THE father name processing logic SHALL replicate the ProcessFatherName behavior from the Legacy_VBA_System

### Requirement 7: Gender Normalization

**User Story:** As a data processor, I want to normalize gender values from various representations, so that I can produce consistent gender codes.

#### Acceptance Criteria

1. WHEN processing gender, THE System SHALL search for a header with exact text "מין\n1=זכר\n2+נקבה" (with line breaks)
2. THE System SHALL normalize line break characters (vbCrLf, vbCr, vbLf) before matching
3. WHEN the gender header is found, THE System SHALL insert a Corrected_Column with header "מין - מתוקן"
4. FOR EACH gender value, THE System SHALL trim and convert to lowercase
5. WHEN the value is empty, THE System SHALL set the normalized value to 1 (male)
6. WHEN the value contains "2", "female", "נ", "אישה", or "בת", THE System SHALL set the normalized value to 2 (female)
7. OTHERWISE, THE System SHALL set the normalized value to 1 (male)
8. THE gender normalization logic SHALL replicate the NormalizeGenderValue behavior from the Legacy_VBA_System

### Requirement 8: Date Field Structure

**User Story:** As a data processor, I want to process date fields with split columns for year, month, and day, so that I can handle the specific date structure used in the source data.

#### Acceptance Criteria

1. WHEN processing dates, THE System SHALL search for "תאריך לידה" (birth date) headers
2. WHEN processing dates, THE System SHALL search for "תאריך כניסה למוסד" (entry date) headers
3. FOR EACH found date header, THE System SHALL search the next row for sub-headers "שנה" (year), "חודש" (month), and "יום" (day)
4. WHEN all three sub-headers are found, THE System SHALL insert 4 Corrected_Columns after the day column
5. THE Corrected_Column headers SHALL be "שנה - מתוקן", "חודש - מתוקן", "יום - מתוקן", "סטטוס תאריך"
6. THE System SHALL format the year, month, and day columns as "0" (number format)
7. THE date field structure logic SHALL replicate the ProcessDateField behavior from the Legacy_VBA_System

### Requirement 9: Date Format Pattern Detection

**User Story:** As a data processor, I want to detect the dominant date format pattern for ambiguous dates, so that I can parse dates correctly.

#### Acceptance Criteria

1. WHEN processing a date column, THE System SHALL sample all date values in the column
2. FOR EACH date value containing "/" or ".", THE System SHALL split by the separator
3. WHEN the first part is greater than 12 and the second part is less than or equal to 12, THE System SHALL count it as DDMM
4. WHEN the second part is greater than 12 and the first part is less than or equal to 12, THE System SHALL count it as MMDD
5. WHEN MMDD count is greater than DDMM count, THE System SHALL use Date_Format_Pattern MMDD
6. OTHERWISE, THE System SHALL use Date_Format_Pattern DDMM
7. THE pattern detection logic SHALL replicate the DetectDominantPattern behavior from the Legacy_VBA_System

### Requirement 10: Date Parsing from Split Columns

**User Story:** As a data processor, I want to parse dates from split year, month, and day columns, so that I can handle dates that are already separated.

#### Acceptance Criteria

1. WHEN year, month, and day values are all non-empty and numeric, THE System SHALL parse from split columns
2. THE System SHALL convert year, month, and day to long integers
3. WHEN the year is less than 100, THE System SHALL expand it to a four-digit year using two-digit year expansion
4. THE System SHALL validate that day is between 1 and 31
5. THE System SHALL validate that month is between 1 and 12
6. THE System SHALL validate that the date exists using DateSerial
7. WHEN validation fails, THE System SHALL set a status text describing the error
8. THE split column parsing logic SHALL replicate the ParseFromSplitColumns behavior from the Legacy_VBA_System

### Requirement 11: Date Parsing from Main Value

**User Story:** As a data processor, I want to parse dates from the main date column when split columns are empty, so that I can handle dates in various formats.

#### Acceptance Criteria

1. WHEN split columns are empty, THE System SHALL parse from the main date value
2. WHEN the value is empty, THE System SHALL set status text to "תא ריק" (empty cell)
3. WHEN the value is a VBA Date type, THE System SHALL extract year, month, and day directly
4. WHEN the value contains only digits (no separators), THE System SHALL parse as numeric date string
5. FOR 8-digit strings, THE System SHALL parse as DDMMYYYY
6. FOR 6-digit strings, THE System SHALL parse as DDMMYY with two-digit year expansion
7. FOR 4-digit strings, THE System SHALL parse as DMYY with two-digit year expansion
8. WHEN the value contains "/" or ".", THE System SHALL parse as separated date string using the dominant pattern
9. THE main value parsing logic SHALL replicate the ParseDateValue behavior from the Legacy_VBA_System

### Requirement 12: Date Business Rule Validation

**User Story:** As a data processor, I want to validate dates against business rules, so that I can identify invalid or suspicious dates.

#### Acceptance Criteria

1. WHEN the field type is entry date and the status text is "תא ריק", THE System SHALL clear the status text (entry date is optional)
2. WHEN the date is not valid, THE System SHALL keep the existing status text
3. WHEN the year is less than 1900, THE System SHALL set status text to "שנה לפני 1900" and mark as invalid
4. WHEN the date is in the future, THE System SHALL set status text to "תאריך לידה עתידי" or "תאריך כניסה עתידי" and mark as invalid
5. WHEN the field type is birth date and the age is greater than 100 years, THE System SHALL set status text to "גיל מעל 100 (X שנים)" but keep the date as valid
6. THE System SHALL calculate age by subtracting birth year from current year, then adjusting if the birthday hasn't occurred yet this year
7. THE business rule validation logic SHALL replicate the ValidateBusinessRules behavior from the Legacy_VBA_System

### Requirement 13: Date Output Writing

**User Story:** As a data processor, I want to write parsed dates to the corrected columns with appropriate formatting, so that I can produce standardized date output.

#### Acceptance Criteria

1. WHEN a date is valid, THE System SHALL write the year, month, and day to the three corrected columns
2. THE System SHALL always write the status text to the fourth corrected column
3. WHEN the status text is not empty and contains "גיל מעל", THE System SHALL apply yellow background (RGB 255, 230, 150) and bold font to the status cell
4. WHEN the status text is not empty and does not contain "גיל מעל", THE System SHALL apply pink background (RGB 255, 200, 200) and bold font to the status cell
5. THE date output writing logic SHALL replicate the WriteDateResult behavior from the Legacy_VBA_System

### Requirement 14: Identifier Field Processing

**User Story:** As a data processor, I want to process Israeli ID and passport fields together, so that I can validate IDs and move invalid IDs to the passport field.

#### Acceptance Criteria

1. WHEN processing identifiers, THE System SHALL search for "מספר זהות" (ID number) headers
2. WHEN processing identifiers, THE System SHALL search for "מספר דרכון" (passport number) headers
3. WHEN both headers are found, THE System SHALL insert 3 Corrected_Columns after the passport column
4. THE Corrected_Column headers SHALL be "ת.ז. - מתוקן", "דרכון - מתוקן", "סטטוס מזהה"
5. THE System SHALL read both ID and passport columns to the same last row (maximum of both columns)
6. THE System SHALL process identifiers using the identifier normalization logic
7. THE System SHALL write the corrected ID, corrected passport, and status to the three corrected columns
8. THE identifier field processing logic SHALL replicate the ProcessIdentifiers behavior from the Legacy_VBA_System

### Requirement 15: Identifier Value Classification

**User Story:** As a data processor, I want to classify identifier values as Israeli ID or passport, so that I can apply the appropriate validation logic.

#### Acceptance Criteria

1. WHEN the ID value is "9999", THE System SHALL treat it as empty
2. WHEN the ID value contains any character that is NOT a digit or dash (including en-dash, em-dash, minus sign), THE System SHALL move it to the passport field
3. THE System SHALL accept multiple dash Unicode characters (hyphen 45, non-breaking hyphen 8209, figure dash 8210, en-dash 8211, em-dash 8212, horizontal bar 8213, minus sign 8722)
4. WHEN the ID value is moved to passport, THE System SHALL set status to "ת.ז. הועברה לדרכון"
5. THE System SHALL extract digits only (removing all dashes) for validation
6. WHEN the digit count is less than 4, THE System SHALL move the ID to passport and set status to "ת.ז. לא תקינה + הועברה לדרכון"
7. WHEN the digit count is greater than 9, THE System SHALL move the ID to passport and set status to "ת.ז. הועברה לדרכון"
8. THE identifier classification logic SHALL replicate the ProcessIDValue behavior from the Legacy_VBA_System

### Requirement 16: Israeli ID Validation

**User Story:** As a data processor, I want to validate Israeli ID numbers with checksum verification, so that I can identify invalid identifiers.

#### Acceptance Criteria

1. WHEN the ID has 4-9 digits, THE System SHALL pad it to 9 digits with leading zeros
2. WHEN the padded ID is "000000000", THE System SHALL mark it as invalid
3. WHEN all 9 digits are identical, THE System SHALL mark it as invalid
4. THE System SHALL calculate the checksum using the Israeli ID algorithm (multiply odd positions by 1, even positions by 2, subtract 9 if result > 9, sum all, check if last digit makes sum divisible by 10)
5. WHEN the checksum is valid, THE System SHALL set status to "ת.ז. תקינה"
6. WHEN the checksum is invalid, THE System SHALL set status to "ת.ז. לא תקינה"
7. WHEN a passport value is also present, THE System SHALL append " + דרכון הוזן" to the status
8. THE ID validation logic SHALL replicate the ValidateChecksum behavior from the Legacy_VBA_System

### Requirement 17: Passport Value Cleaning

**User Story:** As a data processor, I want to clean passport values by removing invalid characters, so that I can produce standardized passport values.

#### Acceptance Criteria

1. THE System SHALL keep digits (0-9) in passport values
2. THE System SHALL keep English letters (A-Z, a-z) in passport values
3. THE System SHALL keep Hebrew letters (Unicode 1488-1514) in passport values
4. THE System SHALL keep dash characters (all dash Unicode variants) in passport values
5. THE System SHALL remove all other characters from passport values
6. THE passport cleaning logic SHALL replicate the CleanPassportValue behavior from the Legacy_VBA_System

### Requirement 18: Identifier Status Messages

**User Story:** As a data processor, I want clear status messages for identifier processing, so that I can understand what happened to each identifier.

#### Acceptance Criteria

1. WHEN the ID is empty and passport is empty, THE System SHALL set status to "חסר מזהים"
2. WHEN the ID is empty and passport is present, THE System SHALL set status to "דרכון הוזן"
3. WHEN the ID is valid and passport is empty, THE System SHALL set status to "ת.ז. תקינה"
4. WHEN the ID is valid and passport is present, THE System SHALL set status to "ת.ז. תקינה + דרכון הוזן"
5. WHEN the ID is invalid and passport is empty, THE System SHALL set status to "ת.ז. לא תקינה"
6. WHEN the ID is invalid and passport is present, THE System SHALL set status to "ת.ז. לא תקינה + דרכון הוזן"
7. WHEN the ID is moved to passport due to invalid format, THE System SHALL set status to "ת.ז. הועברה לדרכון"
8. WHEN the ID is moved to passport due to too few digits, THE System SHALL set status to "ת.ז. לא תקינה + הועברה לדרכון"
9. THE status message logic SHALL replicate the NormalizeIdentifiers behavior from the Legacy_VBA_System

### Requirement 19: Processing Pipeline

**User Story:** As a developer, I want a clearly defined processing pipeline, so that I can understand the system flow and maintain consistency.

#### Acceptance Criteria

1. THE System SHALL process workbooks in the following order:
   - Step 1: Load workbook
   - Step 2: For each worksheet:
     - Step 2a: Process names (first, last, father)
     - Step 2b: Process gender
     - Step 2c: Process dates (birth, entry)
     - Step 2d: Process identifiers (ID and passport)
   - Step 3: Save workbook
2. THE System SHALL disable screen updating, events, and automatic calculation during processing
3. THE System SHALL re-enable screen updating, events, and automatic calculation after processing
4. THE pipeline SHALL replicate the NormalizeAllWorksheets behavior from the Legacy_VBA_System

### Requirement 20: Corrected Column Tracking

**User Story:** As a developer, I want to track the positions of corrected columns, so that downstream systems can locate the normalized data.

#### Acceptance Criteria

1. WHEN a Corrected_Column is created, THE System SHALL record the worksheet name, field key, and column number
2. THE System SHALL use the following field keys: ShemPrati, ShemMishpaha, ShemHaAv, Min, ShnatLida, HodeshLida, YomLida, shnatknisa, Hodeshknisa, YomKnisa, MisparZehut, Darkon
3. THE System SHALL provide a method to retrieve the column number for a given worksheet and field key
4. THE corrected column tracking SHALL support the same field keys used in the Legacy_VBA_System

### Requirement 21: Extensibility and Modularity

**User Story:** As a developer, I want a modular architecture with clear interfaces, so that I can add new field processors easily.

#### Acceptance Criteria

1. THE System SHALL implement separate classes for ExcelReader, ExcelWriter, and FieldProcessor
2. THE System SHALL implement specialized Engines (NameEngine, DateEngine, GenderEngine, IdentifierEngine) with pure business logic
3. THE System SHALL separate Excel I/O operations from business logic
4. THE System SHALL use the Template Method pattern for field processing
5. THE System SHALL organize code into layers (io_layer, engines, processing)
6. THE architecture SHALL replicate the separation of concerns from the Legacy_VBA_System

### Requirement 22: Type Safety and Documentation

**User Story:** As a developer, I want type hints and comprehensive documentation, so that I can understand and maintain the codebase.

#### Acceptance Criteria

1. THE System SHALL include Python type hints for all function parameters and return values
2. THE System SHALL include docstrings for all public classes and methods
3. THE System SHALL follow Google or NumPy docstring format
4. THE System SHALL include inline comments for complex logic
5. THE System SHALL include a README file with setup and usage instructions
6. THE documentation SHALL reference the Legacy_VBA_System behavior where applicable

### Requirement 23: Testing Requirements

**User Story:** As a developer, I want comprehensive tests, so that I can verify correctness and prevent regressions.

#### Acceptance Criteria

1. THE System SHALL include unit tests for all Engine classes
2. THE System SHALL include unit tests for ExcelReader and ExcelWriter
3. THE System SHALL include integration tests processing sample Excel files
4. THE System SHALL include tests for error handling scenarios
5. THE System SHALL achieve minimum 80% code coverage
6. THE System SHALL include property-based tests for Israeli_ID Checksum validation
7. THE System SHALL include round-trip tests for date parsing
8. THE System SHALL use pytest as the testing framework
9. THE System SHALL include tests comparing Python output to Legacy_VBA_System output for the same input data

### Requirement 24: Performance Requirements

**User Story:** As a data processor, I want to process large files efficiently, so that I can handle production workloads.

#### Acceptance Criteria

1. WHEN processing a Workbook with 10,000 rows, THE System SHALL complete within 60 seconds
2. THE System SHALL process data in memory using array operations
3. THE System SHALL minimize cell-by-cell operations
4. THE performance SHALL be comparable to or better than the Legacy_VBA_System

### Requirement 25: Logging and Observability

**User Story:** As a system administrator, I want detailed logs of processing activities, so that I can troubleshoot issues and monitor performance.

#### Acceptance Criteria

1. THE System SHALL write logs to a file in the output directory
2. THE System SHALL include log level, timestamp, and message in each log entry
3. THE System SHALL log the start and completion of each processing phase
4. THE System SHALL log summary statistics for each Worksheet processed
5. THE System SHALL support log rotation to prevent unbounded file growth
6. THE System SHALL output logs to console when running in CLI mode

### Requirement 26: Data Privacy and Security

**User Story:** As a compliance officer, I want the system to handle personal data securely, so that we meet privacy regulations.

#### Acceptance Criteria

1. THE System SHALL not transmit data over networks
2. THE System SHALL not store credentials or sensitive configuration in code
3. THE System SHALL support file permissions on output files
4. THE System SHALL include a warning in documentation about handling personal data
5. THE System SHALL not include sample data with real personal information in the repository
