# Implementation Plan: JSON-Based Excel Data Normalization Pipeline

## Overview

This implementation plan transforms the Excel data normalization system to use JSON as the internal data representation. The system will extract Excel data to JSON, apply normalization engines to JSON rows, and export corrected JSON back to Excel. This approach decouples Excel complexity from normalization logic.

## Architecture Changes

**Current**: Excel → Direct Processing → Excel
**New**: Excel → JSON → Normalization → JSON → Excel

## Existing Code to Reuse

- ✅ ExcelReader with header detection (`detect_columns`, `detect_table_region`)
- ✅ Header detection logic (`_normalize_text`, `_match_field`, `_score_header_row`)
- ✅ Multi-row header support (`_score_subheader_row`, `_detect_date_subcolumns`)
- ✅ Merged cell handling (`_is_merged_cell`, `_get_merged_cell_range`)
- ✅ NameEngine, GenderEngine, DateEngine, IdentifierEngine (NO CHANGES)

## Tasks

### Phase 1: Data Models and JSON Schema

- [ ] 1. Create JSON data models
  - [x] 1.1 Create `JsonRow` type alias
    - Define as `Dict[str, Any]`
    - Document field naming convention (field_name, field_name_corrected)
    - Add type hints for common field types
    - _Requirements: 10.2, 14.3-14.4_

  - [x] 1.2 Create `SheetDataset` dataclass
    - Add fields: sheet_name, header_row, header_rows_count, field_names, rows, metadata
    - Add validation methods
    - Add helper methods (get_field_names, get_row_count)
    - _Requirements: 11.2-11.5, 19.1-19.2_

  - [x] 1.3 Create `WorkbookDataset` dataclass
    - Add fields: source_file, sheets, metadata
    - Add methods to access sheets by name
    - Add validation for multi-sheet datasets
    - _Requirements: 16.1-16.4_

  - [x] 1.4 Document JSON schema
    - Create JSON schema file for SheetDataset
    - Create JSON schema file for JsonRow
    - Add schema validation utilities
    - Document corrected field naming convention
    - _Requirements: 19.1-19.5_

### Phase 2: Excel to JSON Extraction

- [ ] 2. Implement ExcelToJsonExtractor class
  - [x] 2.1 Create ExcelToJsonExtractor class structure
    - Initialize with ExcelReader dependency
    - Add configuration options (skip empty rows, handle formulas)
    - _Requirements: 10.1, 20.2_

  - [x] 2.2 Implement `extract_row_to_json` method
    - Accept worksheet, row number, column mapping
    - Read cell values for each field in mapping
    - Handle empty cells (store None or empty string)
    - Handle formula cells (extract calculated value)
    - Return JsonRow dictionary
    - _Requirements: 10.2-10.7_

  - [x] 2.3 Implement `extract_sheet_to_json` method
    - Use ExcelReader to detect headers and table region
    - Get column mapping from detect_columns
    - Extract all data rows using extract_row_to_json
    - Create SheetDataset with metadata
    - Handle multi-row headers (store header_rows_count)
    - _Requirements: 10.1-10.5, 11.1-11.5_

  - [x] 2.4 Implement `extract_workbook_to_json` method
    - Open workbook using openpyxl
    - Iterate through all worksheets
    - Call extract_sheet_to_json for each sheet
    - Skip sheets with no valid headers
    - Create WorkbookDataset with all sheets
    - Add workbook-level metadata
    - _Requirements: 16.1-16.6_

  - [x] 2.5 Add error handling for extraction
    - Handle missing headers gracefully
    - Handle invalid cell values
    - Handle merged cells
    - Log warnings for skipped sheets
    - _Requirements: 18.1-18.4_

### Phase 3: Normalization Pipeline

- [ ] 3. Implement NormalizationPipeline class
  - [x] 3.1 Create NormalizationPipeline class structure
    - Initialize with engine dependencies (NameEngine, GenderEngine, etc.)
    - Add configuration for which engines to apply
    - _Requirements: 12.1, 17.6_

  - [x] 3.2 Implement `normalize_row` method
    - Accept JsonRow as input
    - Create copy of row for modifications
    - Call apply_name_normalization
    - Call apply_gender_normalization
    - Call apply_date_normalization
    - Call apply_identifier_normalization
    - Return row with corrected fields
    - _Requirements: 12.2, 13.2-13.5_

  - [x] 3.3 Implement `apply_name_normalization` method
    - Check for first_name, last_name, father_name fields
    - Call NameEngine.normalize_name for each field
    - Store results in field_name_corrected keys
    - Handle None/empty values
    - Handle engine exceptions
    - _Requirements: 12.3, 12.8, 14.1-14.5_

  - [x] 3.4 Implement `apply_gender_normalization` method
    - Check for gender field
    - Call GenderEngine.normalize_gender
    - Store result in gender_corrected key
    - Handle None/empty values
    - Handle engine exceptions
    - _Requirements: 12.4, 12.8, 14.1-14.5_

  - [x] 3.5 Implement `apply_date_normalization` method
    - Check for birth_date or birth_year/month/day fields
    - Check for entry_date or entry_year/month/day fields
    - Call DateEngine.normalize_date with appropriate parameters
    - Store results in corrected keys
    - Handle single vs split date fields
    - Handle None/empty values
    - Handle engine exceptions
    - _Requirements: 12.5, 12.8, 14.1-14.5_

  - [x] 3.6 Implement `apply_identifier_normalization` method
    - Check for id_number and passport fields
    - Call IdentifierEngine.normalize_id for each field
    - Store results in corrected keys
    - Handle None/empty values
    - Handle engine exceptions
    - _Requirements: 12.6, 12.8, 14.1-14.5_

  - [x] 3.7 Implement `normalize_dataset` method
    - Accept SheetDataset as input
    - Create copy of dataset
    - Iterate through all rows
    - Call normalize_row for each row
    - Update metadata with normalization info
    - Return corrected dataset
    - _Requirements: 12.1-12.2, 13.1-13.7_

  - [x] 3.8 Add error handling for normalization
    - Handle engine failures gracefully
    - Store original value if engine fails
    - Log errors with row number and field name
    - Continue processing other fields/rows
    - _Requirements: 18.1-18.4_

### Phase 4: JSON to Excel Export

- [ ] 4. Implement JsonToExcelWriter class
  - [x] 4.1 Create JsonToExcelWriter class structure
    - Add configuration options (column widths, formatting)
    - _Requirements: 15.1, 20.4_

  - [x] 4.2 Implement `create_header_row` method
    - Accept worksheet and field names list
    - Write original field names in columns
    - Write corrected field names (field_name_corrected) in adjacent columns
    - Apply header formatting (bold, background color)
    - _Requirements: 15.4-15.6_

  - [x] 4.3 Implement `write_json_row` method
    - Accept worksheet, row number, JsonRow, field names
    - Write original values to columns
    - Write corrected values to adjacent columns
    - Handle None values (write empty cell)
    - _Requirements: 15.7_

  - [x] 4.4 Implement `write_dataset_to_excel` method
    - Create new workbook
    - Create worksheet with sheet name
    - Call create_header_row
    - Iterate through rows and call write_json_row
    - Save workbook to output path
    - _Requirements: 15.1-15.8_

  - [x] 4.5 Implement `write_workbook_to_excel` method
    - Create new workbook
    - Iterate through all sheet datasets
    - Create worksheet for each sheet
    - Call write_dataset_to_excel logic for each sheet
    - Save workbook to output path
    - _Requirements: 15.1-15.8, 16.3-16.4_

  - [x] 4.6 Add error handling for export
    - Handle file write failures
    - Validate dataset structure before export
    - Clean up partial files on error
    - _Requirements: 18.1-18.4_

### Phase 5: Integration and Orchestration

- [ ] 5. Update orchestrator to use JSON pipeline
  - [x] 5.1 Create new orchestration method `process_workbook_json`
    - Accept input Excel path and output Excel path
    - Create ExcelToJsonExtractor instance
    - Extract workbook to JSON
    - Create NormalizationPipeline instance
    - Normalize all sheet datasets
    - Create JsonToExcelWriter instance
    - Write corrected datasets to output Excel
    - _Requirements: 20.1-20.6_

  - [x] 5.2 Add support for exporting raw JSON
    - Add method to export RawJsonDataset to JSON file
    - Add method to export CorrectedJsonDataset to JSON file
    - Support both single sheet and workbook export
    - _Requirements: 11.6_

  - [ ] 5.3 Add CLI support for JSON pipeline
    - Add command line option to use JSON pipeline
    - Add option to export intermediate JSON files
    - Add option to import JSON and export to Excel
    - Update help text and documentation
    - _Requirements: 20.1-20.6_

  - [ ] 5.4 Maintain backward compatibility
    - Keep existing direct Excel processing as fallback
    - Add feature flag to switch between pipelines
    - Ensure existing tests still pass
    - _Requirements: 17.1-17.7_

### Phase 6: Testing

- [ ] 6. Add unit tests for new components
  - [ ] 6.1 Test ExcelToJsonExtractor
    - Test extract_row_to_json with various cell types
    - Test extract_sheet_to_json with single-row headers
    - Test extract_sheet_to_json with multi-row headers
    - Test extract_workbook_to_json with multiple sheets
    - Test handling of empty cells
    - Test handling of merged cells
    - Test handling of formula cells
    - _Requirements: 10.1-10.7_

  - [ ] 6.2 Test NormalizationPipeline
    - Test normalize_row with all field types
    - Test apply_name_normalization
    - Test apply_gender_normalization
    - Test apply_date_normalization (single and split)
    - Test apply_identifier_normalization
    - Test normalize_dataset
    - Test error handling for engine failures
    - Test handling of missing fields
    - _Requirements: 12.1-12.8, 13.1-13.7_

  - [ ] 6.3 Test JsonToExcelWriter
    - Test create_header_row
    - Test write_json_row
    - Test write_dataset_to_excel
    - Test write_workbook_to_excel
    - Test handling of None values
    - Test file creation and saving
    - _Requirements: 15.1-15.8_

  - [ ] 6.4 Test data models
    - Test SheetDataset creation and validation
    - Test WorkbookDataset creation and validation
    - Test JsonRow structure
    - Test metadata handling
    - _Requirements: 11.1-11.5, 16.1-16.4_

- [ ] 7. Add integration tests
  - [ ] 7.1 Test end-to-end pipeline
    - Test Excel → JSON → Normalize → Excel flow
    - Test with single worksheet
    - Test with multiple worksheets
    - Test with various header structures
    - Verify original values preserved
    - Verify corrected values generated
    - _Requirements: 14.1-14.5, 20.1-20.6_

  - [ ] 7.2 Test engine compatibility
    - Verify NameEngine works with JSON values
    - Verify GenderEngine works with JSON values
    - Verify DateEngine works with JSON values
    - Verify IdentifierEngine works with JSON values
    - Test all field types end-to-end
    - _Requirements: 17.1-17.7_

  - [ ] 7.3 Test with real-world Excel files
    - Test with files from different institutions
    - Test with complex header structures
    - Test with merged cells and multi-row headers
    - Test with large datasets (1000+ rows)
    - Verify performance targets met
    - _Requirements: 16.1-16.6_

  - [ ] 7.4 Test error scenarios
    - Test with missing headers
    - Test with invalid cell values
    - Test with engine failures
    - Test with file write failures
    - Verify graceful error handling
    - _Requirements: 18.1-18.4_

- [ ] 8. Add property-based tests
  - [ ] 8.1 Property: Data preservation
    - Original values must never be modified
    - Test with random JSON rows
    - **Validates: Requirements 14.1-14.5**

  - [ ] 8.2 Property: Corrected field creation
    - Every field must have corresponding corrected field
    - Test with random JSON rows
    - **Validates: Requirements 13.2-13.5**

  - [ ] 8.3 Property: Round-trip consistency
    - Excel → JSON → Excel should preserve original data
    - Test with random Excel files
    - **Validates: Requirements 10.1-10.7, 15.1-15.8**

  - [ ] 8.4 Property: Field naming convention
    - Corrected fields must follow naming convention
    - Test with random field names
    - **Validates: Requirements 13.3, 19.4**

### Phase 7: Documentation and Cleanup

- [ ] 9. Update documentation
  - [ ] 9.1 Update README with JSON pipeline information
    - Explain new architecture
    - Provide usage examples
    - Document JSON schema
    - _Requirements: 19.1-19.5_

  - [ ] 9.2 Add code documentation
    - Add docstrings to all new classes and methods
    - Add type hints throughout
    - Add usage examples in docstrings
    - _Requirements: 20.1-20.6_

  - [ ] 9.3 Create migration guide
    - Document changes from old to new pipeline
    - Provide migration examples
    - Document backward compatibility
    - _Requirements: 17.1-17.7_

  - [ ] 9.4 Update API documentation
    - Document ExcelToJsonExtractor API
    - Document NormalizationPipeline API
    - Document JsonToExcelWriter API
    - Document data models
    - _Requirements: 19.1-19.5_

- [ ] 10. Final validation and cleanup
  - [ ] 10.1 Run all tests
    - Run unit tests (target: 100% pass)
    - Run integration tests (target: 100% pass)
    - Run property-based tests (target: 100% pass)
    - Verify test coverage (target: >80%)
    - _Requirements: All_

  - [ ] 10.2 Performance validation
    - Test with 1000 row dataset (target: <5 seconds)
    - Test with 10,000 row dataset (target: <30 seconds)
    - Test with multiple worksheets (target: <10 seconds per sheet)
    - Profile and optimize bottlenecks
    - _Requirements: 20.1-20.6_

  - [ ] 10.3 Code review and refactoring
    - Review all new code for quality
    - Refactor duplicated code
    - Ensure consistent naming conventions
    - Ensure proper error handling
    - _Requirements: 20.1-20.6_

  - [ ] 10.4 Verify backward compatibility
    - Run existing tests with new code
    - Verify engines unchanged
    - Verify ExcelReader unchanged
    - Verify no breaking changes
    - _Requirements: 17.1-17.7_

## Notes

- **Reuse existing code**: ExcelReader, header detection, engines remain unchanged
- **Clean separation**: IO, extraction, normalization, export are separate components
- **Non-destructive**: Original values always preserved alongside corrected values
- **Engine compatibility**: Existing engines work without modification
- **Testability**: JSON-based approach makes testing easier
- **Maintainability**: Clear interfaces and separation of concerns
- Each task references specific requirements for traceability
- Property tests validate universal correctness properties
- Integration tests validate end-to-end workflows

## Implementation Order

1. **Phase 1**: Data models (foundation for everything else)
2. **Phase 2**: Extraction (convert Excel to JSON)
3. **Phase 3**: Normalization (apply engines to JSON)
4. **Phase 4**: Export (convert JSON back to Excel)
5. **Phase 5**: Integration (wire everything together)
6. **Phase 6**: Testing (comprehensive test coverage)
7. **Phase 7**: Documentation (finalize and document)

## Success Criteria

- All 121 existing tests continue to pass
- All new tests pass (target: 100+ new tests)
- Performance targets met (<5 seconds for 1000 rows)
- Engines require zero changes
- ExcelReader requires zero changes
- Clean separation between IO and business logic
- Original values always preserved
- JSON schema documented and validated

