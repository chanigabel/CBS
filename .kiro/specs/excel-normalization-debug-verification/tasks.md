# Implementation Plan

- [x] 1. Write bug condition exploration test
  - **Property 1: Fault Condition** - Debugging Artifacts Missing
  - **CRITICAL**: This test MUST FAIL on unfixed code - failure confirms the bug exists
  - **DO NOT attempt to fix the test or the code when it fails**
  - **NOTE**: This test encodes the expected behavior - it will validate the fix when it passes after implementation
  - **GOAL**: Surface counterexamples that demonstrate the bug exists
  - **Scoped PBT Approach**: Scope the property to concrete failing case - run pipeline on sample Excel file and verify debugging artifacts are missing
  - Test that when process_workbook_json executes, raw_dataset.json does NOT exist, normalized_dataset.json does NOT exist, and console output does NOT contain engine statistics (from Fault Condition in design)
  - The test assertions should match the Expected Behavior Properties from design: file_exists("raw_dataset.json"), file_exists("normalized_dataset.json"), console_output_contains_engine_statistics
  - Run test on UNFIXED code
  - **EXPECTED OUTCOME**: Test FAILS (this is correct - it proves the bug exists)
  - Document counterexamples found: which artifacts are missing, what console output shows instead
  - Mark task complete when test is written, run, and failure is documented
  - _Requirements: 1.1, 1.2, 1.3, 1.4_

- [x] 2. Write preservation property tests (BEFORE implementing fix)
  - **Property 2: Preservation** - Normalization Behavior Unchanged
  - **IMPORTANT**: Follow observation-first methodology
  - Observe behavior on UNFIXED code: run pipeline on sample Excel file and capture the output Excel file content
  - Write property-based tests capturing observed behavior patterns from Preservation Requirements: Excel output identical, engine execution order preserved, corrected fields created with "_corrected" suffix, log file content unchanged
  - Property-based testing generates many test cases for stronger guarantees across different Excel file structures
  - Run tests on UNFIXED code
  - **EXPECTED OUTCOME**: Tests PASS (this confirms baseline behavior to preserve)
  - Mark task complete when tests are written, run, and passing on unfixed code
  - _Requirements: 3.1, 3.2, 3.3, 3.4_

- [x] 3. Fix for missing debugging artifacts in normalization pipeline

  - [x] 3.1 Add raw JSON export after extraction
    - Import JsonExporter class in orchestrator.py
    - After extractor.extract_workbook_to_json() completes (line ~230), create JsonExporter instance with indent=2, ensure_ascii=False
    - Call exporter.export_workbook_to_json(workbook_dataset, "raw_dataset.json")
    - Log export completion
    - Verify that JSON export does not modify workbook_dataset (read-only operation)
    - _Bug_Condition: isBugCondition(pipeline_execution) where NOT file_exists("raw_dataset.json")_
    - _Expected_Behavior: file_exists("raw_dataset.json") AND raw_json_represents_exact_extraction(result)_
    - _Preservation: Excel output unchanged, engine execution order preserved_
    - _Requirements: 2.1, 2.5_

  - [x] 3.2 Add normalized JSON export after normalization
    - After all sheets are normalized (line ~290), use same JsonExporter instance
    - Call exporter.export_workbook_to_json(workbook_dataset, "normalized_dataset.json")
    - Log export completion
    - Verify that JSON export does not modify workbook_dataset (read-only operation)
    - _Bug_Condition: isBugCondition(pipeline_execution) where NOT file_exists("normalized_dataset.json")_
    - _Expected_Behavior: file_exists("normalized_dataset.json") AND normalized_json_contains_all_corrected_fields(result)_
    - _Preservation: Excel output unchanged, engine execution order preserved_
    - _Requirements: 2.3, 2.5_

  - [x] 3.3 Add engine statistics tracking
    - After each sheet normalization (line ~270), extract statistics from normalized_sheet.get_metadata("normalization_statistics", {})
    - Accumulate counts across all sheets for each engine: NameEngine, GenderEngine, DateEngine, IdentifierEngine
    - Track total rows processed and total modifications per engine
    - _Bug_Condition: isBugCondition(pipeline_execution) where NOT console_output_contains_engine_statistics_
    - _Expected_Behavior: console_output_contains_engine_statistics(result) showing all four engines_
    - _Preservation: Excel output unchanged, engine execution order preserved_
    - _Requirements: 2.2, 2.4_

  - [x] 3.4 Add console summary report
    - After normalized JSON export, print summary to console
    - Print total rows processed across all sheets
    - Print number of values modified by NameEngine
    - Print number of values modified by GenderEngine
    - Print number of values modified by DateEngine
    - Print number of values modified by IdentifierEngine
    - Print verification guidance suggesting which component to investigate
    - _Bug_Condition: isBugCondition(pipeline_execution) where NOT console_output_contains_engine_statistics_
    - _Expected_Behavior: console_output_contains_engine_statistics(result) with verification guidance_
    - _Preservation: Existing log file output unchanged_
    - _Requirements: 2.4, 2.5_

  - [x] 3.5 Verify bug condition exploration test now passes
    - **Property 1: Expected Behavior** - Debugging Artifacts Present
    - **IMPORTANT**: Re-run the SAME test from task 1 - do NOT write a new test
    - The test from task 1 encodes the expected behavior
    - When this test passes, it confirms the expected behavior is satisfied
    - Run bug condition exploration test from step 1
    - **EXPECTED OUTCOME**: Test PASSES (confirms bug is fixed)
    - Verify that raw_dataset.json exists, normalized_dataset.json exists, and console output contains all engine statistics
    - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.5_

  - [x] 3.6 Verify preservation tests still pass
    - **Property 2: Preservation** - Normalization Behavior Unchanged
    - **IMPORTANT**: Re-run the SAME tests from task 2 - do NOT write new tests
    - Run preservation property tests from step 2
    - **EXPECTED OUTCOME**: Tests PASS (confirms no regressions)
    - Confirm that Excel output is identical, engine execution order preserved, corrected fields created correctly, and log file unchanged
    - _Requirements: 3.1, 3.2, 3.3, 3.4_

- [x] 4. Checkpoint - Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise
  - Verify debugging artifacts enable defect isolation: raw_dataset.json can be compared against input Excel, normalized_dataset.json can be compared against raw_dataset.json, output Excel can be compared against normalized_dataset.json
