# Implementation Plan: Excel Data Normalization Python

## Overview

This implementation plan breaks down the Python-based Excel data normalization system into discrete coding tasks. The system replicates the exact behavior of a legacy VBA implementation, processing Excel workbooks to normalize person records (names, gender, dates, identifiers) with strict behavioral equivalence.

The implementation follows a layered architecture with strict separation of concerns: only ExcelReader/ExcelWriter interact with openpyxl, while all business logic operates on plain Python data structures.

## Tasks

- [x] 1. Set up project structure and core data types
  - Create directory structure (src/excel_normalization with io_layer, processing, engines subdirectories)
  - Create __init__.py files for all packages
  - Implement data_types.py with all dataclasses and enums (ColumnHeaderInfo, DateParseResult, IdentifierResult, Language, FatherNamePattern, DateFormatPattern, DateFieldType, FieldKey)
  - Set up pyproject.toml with dependencies (openpyxl 3.1+, pytest 7.0+, pytest-cov, hypothesis)
  - Create requirements.txt
  - _Requirements: 21.1, 21.5, 22.1_

- [ ]* 1.1 Write unit tests for data types
  - Test dataclass instantiation and field access
  - Test enum values and string representations
  - _Requirements: 23.1_

- [x] 2. Implement I/O layer (ExcelReader and ExcelWriter)
  - [x] 2.1 Implement ExcelReader class
    - Implement find_header method with exact text matching (xlPart equivalent) and optional line break normalization
    - Implement read_column_array method to read column data as list
    - Implement read_cell_value method for single cell access
    - Implement get_last_row method to find last non-empty row
    - _Requirements: 2.1, 2.2, 2.3, 2.7_
  
  - [ ]* 2.2 Write unit tests for ExcelReader
    - Test header finding with multiple variants
    - Test column array reading
    - Test last row detection
    - Test line break normalization in headers
    - _Requirements: 23.2_
  
  - [x] 2.3 Implement ExcelWriter class
    - Implement prepare_output_column method to insert new column with header
    - Implement write_column_array method to write list to column
    - Implement write_cell_value method for single cell writes
    - Implement format_cell method for background color, bold, and number format
    - Implement highlight_changed_cells method to apply pink highlight where values differ
    - _Requirements: 4.3, 4.7, 13.1, 13.3, 13.4_
  
  - [ ]* 2.4 Write unit tests for ExcelWriter
    - Test column insertion and header setting
    - Test array writing
    - Test cell formatting (colors, bold, number format)
    - Test change highlighting logic
    - _Requirements: 23.2_

- [x] 3. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 4. Implement business logic engines (pure functions)
  - [x] 4.1 Implement TextProcessor class
    - Implement detect_language_dominance method (count Hebrew vs English letters)
    - Implement remove_diacritics method using character code mappings
    - Implement fix_hebrew_final_letters method to add spaces after final letters
    - Implement collapse_spaces method to replace multiple spaces with single space
    - Implement clean_text method orchestrating all text normalization steps
    - _Requirements: 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 3.8, 3.9_
  
  - [ ]* 4.2 Write unit tests for TextProcessor
    - Test language detection with various Hebrew/English mixes
    - Test diacritic removal with accented characters
    - Test Hebrew final letter spacing
    - Test space collapsing
    - Test end-to-end text cleaning
    - _Requirements: 23.1_
  
  - [ ]* 4.3 Write property test for TextProcessor
    - **Property 3: Text Normalization Consistency**
    - **Validates: Requirements 3.1, 3.3, 3.4, 3.5, 3.6, 3.7**
  
  - [x] 4.4 Implement NameEngine class
    - Implement normalize_name method using TextProcessor
    - Implement remove_last_name_from_father method with pattern-based removal (RemoveFirst, RemoveLast, None)
    - _Requirements: 4.5, 6.5, 6.6, 6.7_
  
  - [ ]* 4.5 Write unit tests for NameEngine
    - Test name normalization with various inputs
    - Test father name last name removal with all patterns
    - Test edge cases (empty strings, single words)
    - _Requirements: 23.1_
  
  - [x] 4.6 Implement GenderEngine class
    - Implement normalize_gender method with female pattern matching and male default
    - Support patterns: "2", "female", "נ", "אישה", "בת" (case-insensitive)
    - Default empty values to 1 (male)
    - _Requirements: 7.4, 7.5, 7.6, 7.7_
  
  - [ ]* 4.7 Write unit tests for GenderEngine
    - Test all female patterns
    - Test male default for non-matching values
    - Test empty value handling
    - _Requirements: 23.1_
  
  - [ ]* 4.8 Write property tests for GenderEngine
    - **Property 8: Gender Female Pattern Recognition**
    - **Property 9: Gender Default to Male**
    - **Validates: Requirements 7.6, 7.7**

- [x] 5. Implement DateEngine class
  - [x] 5.1 Implement date parsing core methods
    - Implement expand_two_digit_year method (if year <= current_year % 100, then 2000 + year, else 1900 + year)
    - Implement parse_from_split_columns method for year/month/day columns with validation
    - Implement parse_numeric_date_string method for 8-digit (DDMMYYYY), 6-digit (DDMMYY), 4-digit (DMYY) formats
    - Implement parse_separated_date_string method for "/" and "." separators using dominant pattern
    - Implement parse_from_main_value method orchestrating all parsing strategies
    - _Requirements: 10.1, 10.2, 10.3, 10.4, 10.5, 10.6, 11.1, 11.2, 11.3, 11.4, 11.5, 11.6, 11.7, 11.8_
  
  - [x] 5.2 Implement date validation methods
    - Implement validate_business_rules method checking year >= 1900, not future, age <= 100
    - Implement calculate_age method for age calculation
    - Set appropriate Hebrew status messages for each validation failure
    - _Requirements: 12.1, 12.2, 12.3, 12.4, 12.5, 12.6_
  
  - [x] 5.3 Implement parse_date orchestration method
    - Combine split column parsing and main value parsing
    - Apply business rule validation
    - Return DateParseResult with year, month, day, is_valid, status_text
    - _Requirements: 10.7, 11.9, 12.7_
  
  - [ ]* 5.4 Write unit tests for DateEngine
    - Test two-digit year expansion with various years
    - Test split column parsing with valid and invalid dates
    - Test all numeric date formats (8-digit, 6-digit, 4-digit)
    - Test separated date parsing with DDMM and MMDD patterns
    - Test business rule validation (pre-1900, future, age > 100)
    - Test empty cell handling
    - _Requirements: 23.1_
  
  - [ ]* 5.5 Write property tests for DateEngine
    - **Property 10: Two-Digit Year Expansion**
    - **Property 11: Date Component Range Validation**
    - **Property 12: Date Existence Validation**
    - **Property 13: Eight-Digit Date Parsing**
    - **Property 14: Six-Digit Date Parsing**
    - **Property 15: Four-Digit Date Parsing**
    - **Property 16: Separated Date Parsing**
    - **Property 17: Pre-1900 Date Rejection**
    - **Property 18: Future Date Rejection**
    - **Property 19: Age Over 100 Warning**
    - **Validates: Requirements 10.3, 10.4, 10.5, 10.6, 11.5, 11.6, 11.7, 11.8, 12.3, 12.4, 12.5**

- [x] 6. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 7. Implement IdentifierEngine class
  - [x] 7.1 Implement ID classification and validation
    - Implement classify_id_value method to detect non-digit characters, handle "9999", check digit count
    - Implement validate_israeli_id method with checksum algorithm (multiply odd positions by 1, even by 2, subtract 9 if > 9, sum, check divisibility by 10)
    - Implement pad_id method to pad 4-9 digit IDs to 9 digits with leading zeros
    - Check for all-zeros and all-identical-digits rejection
    - _Requirements: 15.1, 15.2, 15.3, 15.4, 15.5, 15.6, 15.7, 16.1, 16.2, 16.3, 16.4_
  
  - [x] 7.2 Implement passport cleaning
    - Implement clean_passport method keeping only digits, English letters, Hebrew letters, and dash variants
    - Support all dash Unicode variants (hyphen, non-breaking hyphen, figure dash, en-dash, em-dash, horizontal bar, minus sign)
    - _Requirements: 17.1, 17.2, 17.3, 17.4, 17.5_
  
  - [x] 7.3 Implement normalize_identifiers orchestration method
    - Process ID and passport values together
    - Move invalid IDs to passport field
    - Generate appropriate Hebrew status messages for all scenarios
    - Return IdentifierResult with corrected_id, corrected_passport, status_text
    - _Requirements: 18.1, 18.2, 18.3, 18.4, 18.5, 18.6, 18.7, 18.8, 18.9_
  
  - [ ]* 7.4 Write unit tests for IdentifierEngine
    - Test ID classification with various invalid formats
    - Test checksum validation with valid and invalid IDs
    - Test ID padding with 4-9 digit IDs
    - Test all-zeros and all-identical rejection
    - Test passport cleaning with various characters
    - Test all status message scenarios
    - Test "9999" special case handling
    - _Requirements: 23.1, 23.6_
  
  - [ ]* 7.5 Write property tests for IdentifierEngine
    - **Property 21: ID Non-Digit Character Rejection**
    - **Property 22: Dash Variant Acceptance**
    - **Property 23: Short ID Rejection**
    - **Property 24: Long ID Rejection**
    - **Property 25: ID Zero Padding**
    - **Property 26: Identical Digit ID Rejection**
    - **Property 27: Israeli ID Checksum Validation**
    - **Property 28: Valid ID Status**
    - **Property 29: Invalid ID Status**
    - **Property 30: ID with Passport Status Suffix**
    - **Property 31: Passport Character Preservation**
    - **Validates: Requirements 15.2, 15.3, 15.6, 15.7, 16.1, 16.3, 16.4, 16.5, 16.6, 16.7, 17.1-17.5**

- [x] 8. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 9. Implement field processors
  - [x] 9.1 Implement FieldProcessor abstract base class
    - Define abstract methods: find_headers, prepare_output_columns, process_data
    - Implement process_field template method calling the three abstract methods in sequence
    - _Requirements: 21.4_
  
  - [x] 9.2 Implement NameFieldProcessor class
    - Implement find_headers to search for "שם פרטי"/"first name", "שם משפחה"/"last name", "שם האב"/"father's name"
    - Implement prepare_output_columns to insert corrected columns with " - מתוקן" suffix
    - Implement detect_father_name_pattern method sampling first 5 rows
    - Implement process_data to normalize names using NameEngine and highlight changes
    - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7, 5.8, 6.1, 6.2, 6.3, 6.4, 6.8, 6.9, 6.10_
  
  - [ ]* 9.3 Write unit tests for NameFieldProcessor
    - Test header finding with all variants
    - Test father name pattern detection with various scenarios
    - Test name normalization integration
    - Test change highlighting
    - _Requirements: 23.1_
  
  - [ ]* 9.4 Write property tests for NameFieldProcessor
    - **Property 2: Header Variant Recognition**
    - **Property 4: Corrected Column Header Format**
    - **Property 5: Change Highlighting**
    - **Property 6: Father Name Last Name Removal**
    - **Validates: Requirements 2.2, 4.4, 4.7, 6.5**
  
  - [x] 9.5 Implement GenderFieldProcessor class
    - Implement find_headers to search for "מין\n1=זכר\n2+נקבה" with line break normalization
    - Implement prepare_output_columns to insert "מין - מתוקן" column
    - Implement process_data to normalize gender using GenderEngine
    - _Requirements: 7.1, 7.2, 7.3, 7.8_
  
  - [ ]* 9.6 Write unit tests for GenderFieldProcessor
    - Test header finding with line break variants
    - Test gender normalization integration
    - _Requirements: 23.1_
  
  - [ ]* 9.7 Write property test for GenderFieldProcessor
    - **Property 7: Line Break Normalization in Headers**
    - **Validates: Requirements 7.2**
  
  - [x] 9.8 Implement DateFieldProcessor class
    - Implement find_headers to search for "תאריך לידה" and "תאריך כניסה למוסד" with sub-headers "שנה", "חודש", "יום"
    - Implement prepare_output_columns to insert 4 columns (year, month, day, status) with "0" number format
    - Implement detect_date_format_pattern method sampling all dates to determine DDMM vs MMDD
    - Implement process_data to parse dates using DateEngine and apply status formatting (yellow for age > 100, pink for errors)
    - _Requirements: 8.1, 8.2, 8.3, 8.4, 8.5, 8.6, 8.7, 9.1, 9.2, 9.3, 9.4, 9.5, 9.6, 9.7, 13.2, 13.5_
  
  - [ ]* 9.9 Write unit tests for DateFieldProcessor
    - Test header finding with sub-headers
    - Test date format pattern detection
    - Test date parsing integration
    - Test status cell formatting (yellow vs pink)
    - _Requirements: 23.1_
  
  - [ ]* 9.10 Write property test for DateFieldProcessor
    - **Property 20: Date Status Formatting**
    - **Validates: Requirements 13.3, 13.4**
  
  - [x] 9.11 Implement IdentifierFieldProcessor class
    - Implement find_headers to search for "מספר זהות" and "מספר דרכון"
    - Implement prepare_output_columns to insert 3 columns (ID, passport, status)
    - Implement process_data to normalize identifiers using IdentifierEngine
    - Read both columns to maximum last row
    - _Requirements: 14.1, 14.2, 14.3, 14.4, 14.5, 14.6, 14.7, 14.8_
  
  - [ ]* 9.12 Write unit tests for IdentifierFieldProcessor
    - Test header finding
    - Test identifier normalization integration
    - Test reading to maximum last row
    - _Requirements: 23.1_

- [x] 10. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

- [x] 11. Implement orchestrator
  - [x] 11.1 Implement NormalizationOrchestrator class
    - Implement normalize_workbook method to load workbook, process all worksheets, save workbook
    - Implement process_worksheet method calling processors in order: names, gender, dates, identifiers
    - Implement corrected column tracking with get_corrected_column method
    - Use FieldKey enum for tracking (ShemPrati, ShemMishpaha, ShemHaAv, Min, ShnatLida, etc.)
    - _Requirements: 1.1, 1.2, 1.3, 1.4, 1.5, 19.1, 19.2, 19.3, 20.1, 20.2, 20.3, 20.4_
  
  - [ ]* 11.2 Write integration tests for NormalizationOrchestrator
    - Create sample Excel workbooks with various data scenarios
    - Test end-to-end processing with all field types
    - Test worksheet iteration
    - Test corrected column tracking
    - Verify original data preservation
    - _Requirements: 23.3, 23.7_
  
  - [ ]* 11.3 Write property test for orchestrator
    - **Property 1: Original Data Preservation**
    - **Validates: Requirements 1.6**

- [x] 12. Implement CLI interface
  - [x] 12.1 Implement cli.py entry point
    - Implement argument parsing for file path
    - Implement error handling for file operations (FileNotFoundError, PermissionError, invalid Excel format)
    - Implement logging configuration (console INFO+, file DEBUG+)
    - Implement main function orchestrating the normalization
    - _Requirements: 25.1, 25.2, 25.3, 25.4, 25.5, 25.6_
  
  - [x] 12.2 Add logging throughout the system
    - Add ERROR logs for file/worksheet failures
    - Add INFO logs for processing milestones (worksheet start/complete, summary stats)
    - Add DEBUG logs for header detection and pattern detection
    - Configure log format with timestamp, level, message
    - _Requirements: 25.1, 25.2, 25.3, 25.4_
  
  - [ ]* 12.3 Write tests for CLI and error handling
    - Test file not found handling
    - Test permission error handling
    - Test invalid Excel format handling
    - Test worksheet-level error recovery
    - Test logging output
    - _Requirements: 23.4_

- [x] 13. Create documentation and setup files
  - [x] 13.1 Create README.md
    - Document system overview and purpose
    - Document installation instructions (pip install -r requirements.txt)
    - Document usage instructions (CLI command examples)
    - Document VBA compatibility notes
    - Include warning about handling personal data
    - _Requirements: 22.5, 22.6, 26.4_
  
  - [x] 13.2 Create setup.py and pyproject.toml
    - Configure package metadata
    - Specify dependencies with versions
    - Configure pytest and coverage settings
    - Configure black, mypy, flake8 settings
    - _Requirements: 22.1_

- [x] 14. Final checkpoint - Ensure all tests pass
  - Run full test suite (unit, property, integration)
  - Verify minimum 80% code coverage
  - Run type checking with mypy
  - Run linting with flake8
  - Run formatting check with black
  - Ensure all tests pass, ask the user if questions arise.

## Notes

- Tasks marked with `*` are optional and can be skipped for faster MVP
- Each task references specific requirements for traceability
- Checkpoints ensure incremental validation at key milestones
- Property tests validate universal correctness properties using Hypothesis
- Unit tests validate specific examples and edge cases
- The implementation strictly follows the layered architecture: only ExcelReader/ExcelWriter touch openpyxl
- All business logic (engines) operates on plain Python data structures
- The system replicates exact VBA behavior - this is a faithful port, not a reimagining
