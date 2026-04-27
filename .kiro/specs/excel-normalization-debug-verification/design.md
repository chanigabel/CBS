# Excel Normalization Debug Verification Bugfix Design

## Overview

The Excel normalization pipeline currently lacks visibility into intermediate processing stages, making it impossible to isolate whether defects originate in extraction, normalization engines, or Excel writing. This bugfix adds temporary debugging infrastructure that exports datasets at critical pipeline stages (raw extraction and post-normalization) and reports engine execution statistics. These debugging artifacts enable systematic verification of each pipeline component without modifying core normalization behavior.

## Glossary

- **Bug_Condition (C)**: The condition where the pipeline lacks debugging visibility - no raw JSON export, no normalized JSON export, and no engine statistics reporting
- **Property (P)**: The desired behavior where debugging artifacts are present - raw_dataset.json exists, normalized_dataset.json exists, and console shows engine statistics
- **Preservation**: Existing normalization behavior that must remain unchanged - same Excel output, same engine execution order, same corrected field creation
- **process_workbook_json**: The method in `orchestrator.py` that orchestrates the JSON-based normalization pipeline
- **ExcelToJsonExtractor**: Component that extracts Excel workbooks to JSON format
- **NormalizationPipeline**: Component that applies all normalization engines to datasets
- **JsonToExcelWriter**: Component that writes normalized datasets back to Excel format
- **Engine Statistics**: Counts of how many values each engine (NameEngine, GenderEngine, DateEngine, IdentifierEngine) modified during normalization

## Bug Details

### Fault Condition

The bug manifests when the normalization pipeline executes without producing debugging artifacts. The `process_workbook_json` method extracts data, normalizes it, and writes Excel output, but does not export intermediate JSON files or report detailed engine statistics, making defect isolation impossible.

**Formal Specification:**
```
FUNCTION isBugCondition(pipeline_execution)
  INPUT: pipeline_execution of type PipelineExecution
  OUTPUT: boolean
  
  RETURN (NOT file_exists("raw_dataset.json")) OR
         (NOT file_exists("normalized_dataset.json")) OR
         (NOT console_output_contains_engine_statistics(pipeline_execution.console_output))
END FUNCTION
```

### Examples

- **Example 1**: User runs normalization on `data.xlsx`, pipeline completes successfully, but no `raw_dataset.json` is created → cannot verify if extraction is correct
- **Example 2**: User runs normalization on `data.xlsx`, pipeline completes successfully, but no `normalized_dataset.json` is created → cannot verify if engines executed correctly
- **Example 3**: User runs normalization on `data.xlsx`, pipeline completes successfully, but console shows no engine statistics → cannot determine which engines modified data or how many values changed
- **Edge Case**: User runs normalization on empty workbook → debugging artifacts should still be created showing 0 rows processed and 0 modifications

## Expected Behavior

### Preservation Requirements

**Unchanged Behaviors:**
- The final normalized Excel output file must remain identical to current behavior
- All normalization engines must continue to execute in the same order: NameEngine, GenderEngine, DateEngine, IdentifierEngine
- Original values must continue to be preserved with corrected fields created using "_corrected" suffix
- Existing log file output must continue to work exactly as before

**Scope:**
All normalization logic, engine execution, and Excel output generation should be completely unaffected by this fix. This includes:
- Engine normalization algorithms and correction logic
- Excel file format and structure
- Column positioning and naming conventions
- Error handling and logging behavior

## Hypothesized Root Cause

Based on the bug description, the pipeline lacks debugging infrastructure:

1. **Missing Raw JSON Export**: The `process_workbook_json` method does not call any JSON export after extraction
   - After `extractor.extract_workbook_to_json()` completes, the raw dataset is not saved
   - Need to add JSON export immediately after extraction step

2. **Missing Normalized JSON Export**: The `process_workbook_json` method does not call any JSON export after normalization
   - After `pipeline.normalize_dataset()` completes for all sheets, the normalized dataset is not saved
   - Need to add JSON export after normalization but before Excel writing

3. **Missing Engine Statistics**: The pipeline does not track or report how many values each engine modified
   - The `NormalizationPipeline` class likely has internal statistics but doesn't expose them
   - Need to extract and report statistics from pipeline after normalization

4. **No Verification Report**: The pipeline does not provide guidance on which component to investigate
   - Need to add console output that summarizes what was processed and suggests next steps

## Correctness Properties

Property 1: Fault Condition - Debugging Artifacts Present

_For any_ pipeline execution where debugging is needed (isBugCondition returns true), the fixed process_workbook_json function SHALL export raw_dataset.json after extraction, export normalized_dataset.json after normalization, and print engine statistics to console showing rows processed and values modified by each engine.

**Validates: Requirements 2.1, 2.2, 2.3, 2.4, 2.5**

Property 2: Preservation - Normalization Behavior Unchanged

_For any_ pipeline execution, the fixed code SHALL produce exactly the same normalized Excel output file as the original code, preserving all normalization logic, engine execution order, corrected field creation, and logging behavior.

**Validates: Requirements 3.1, 3.2, 3.3, 3.4**

## Fix Implementation

### Changes Required

Assuming our root cause analysis is correct:

**File**: `src/excel_normalization/orchestrator.py`

**Function**: `process_workbook_json`

**Specific Changes**:

1. **Add Raw JSON Export**: After extraction completes (line ~230), add JSON export
   - Import `JsonExporter` class
   - Create exporter instance with `indent=2, ensure_ascii=False`
   - Call `exporter.export_workbook_to_json(workbook_dataset, "raw_dataset.json")`
   - Log export completion

2. **Add Normalized JSON Export**: After all sheets are normalized (line ~290), add JSON export
   - Use same `JsonExporter` instance
   - Call `exporter.export_workbook_to_json(workbook_dataset, "normalized_dataset.json")`
   - Log export completion

3. **Add Engine Statistics Tracking**: After each sheet normalization (line ~270), collect statistics
   - Extract statistics from `normalized_sheet.get_metadata("normalization_statistics", {})`
   - Accumulate counts across all sheets for each engine
   - Track: total rows, total modifications per engine

4. **Add Console Summary Report**: After normalized JSON export, print summary
   - Print total rows processed across all sheets
   - Print number of values modified by NameEngine
   - Print number of values modified by GenderEngine
   - Print number of values modified by DateEngine
   - Print number of values modified by IdentifierEngine
   - Print verification guidance suggesting which component to investigate

5. **Ensure Non-Destructive**: Verify that JSON exports do not modify the dataset
   - JSON export should be read-only operation
   - Verify that `workbook_dataset` is not mutated by export calls

## Testing Strategy

### Validation Approach

The testing strategy follows a two-phase approach: first, surface counterexamples that demonstrate the missing debugging artifacts on unfixed code, then verify the fix produces all required artifacts and preserves existing behavior.

### Exploratory Fault Condition Checking

**Goal**: Surface counterexamples that demonstrate the bug BEFORE implementing the fix. Confirm that debugging artifacts are missing and that we cannot isolate defects.

**Test Plan**: Run the normalization pipeline on a sample Excel file using the UNFIXED code. Verify that `raw_dataset.json` and `normalized_dataset.json` are NOT created, and that console output does NOT contain engine statistics.

**Test Cases**:
1. **Missing Raw JSON Test**: Run pipeline on sample Excel file, verify `raw_dataset.json` does not exist (will fail on unfixed code)
2. **Missing Normalized JSON Test**: Run pipeline on sample Excel file, verify `normalized_dataset.json` does not exist (will fail on unfixed code)
3. **Missing Statistics Test**: Run pipeline on sample Excel file, capture console output, verify it does not contain "NameEngine modified:" (will fail on unfixed code)
4. **Empty Workbook Test**: Run pipeline on empty Excel file, verify no debugging artifacts are created (may fail on unfixed code)

**Expected Counterexamples**:
- No JSON files are created in the working directory
- Console output shows only high-level progress messages without engine-specific statistics
- Possible causes: no JSON export calls, no statistics extraction, no console reporting

### Fix Checking

**Goal**: Verify that for all inputs where the bug condition holds, the fixed function produces the expected debugging artifacts.

**Pseudocode:**
```
FOR ALL pipeline_execution WHERE isBugCondition(pipeline_execution) DO
  result := process_workbook_json_fixed(pipeline_execution.input_excel, pipeline_execution.output_excel)
  ASSERT file_exists("raw_dataset.json")
  ASSERT file_exists("normalized_dataset.json")
  ASSERT console_output_contains("NameEngine modified:")
  ASSERT console_output_contains("GenderEngine modified:")
  ASSERT console_output_contains("DateEngine modified:")
  ASSERT console_output_contains("IdentifierEngine modified:")
END FOR
```

### Preservation Checking

**Goal**: Verify that for all inputs, the fixed function produces the same Excel output as the original function.

**Pseudocode:**
```
FOR ALL pipeline_execution DO
  original_excel := process_workbook_json_original(pipeline_execution.input_excel, "output_original.xlsx")
  fixed_excel := process_workbook_json_fixed(pipeline_execution.input_excel, "output_fixed.xlsx")
  ASSERT excel_files_identical(original_excel, fixed_excel)
END FOR
```

**Testing Approach**: Property-based testing is recommended for preservation checking because:
- It generates many test cases automatically across different Excel file structures
- It catches edge cases that manual unit tests might miss (empty sheets, different column orders, missing fields)
- It provides strong guarantees that normalization behavior is unchanged for all inputs

**Test Plan**: Run the UNFIXED code on sample Excel files to capture expected Excel output, then write property-based tests that verify the FIXED code produces identical Excel output.

**Test Cases**:
1. **Excel Output Preservation**: Run both unfixed and fixed code on same input, verify output Excel files are byte-identical
2. **Engine Execution Order Preservation**: Verify that corrected fields appear in same columns in both versions
3. **Logging Preservation**: Verify that log file content is identical (excluding new debug messages)
4. **Error Handling Preservation**: Verify that invalid inputs produce same error messages

### Unit Tests

- Test JSON export after extraction produces valid JSON with correct structure
- Test JSON export after normalization includes corrected fields
- Test engine statistics extraction from normalized sheet metadata
- Test console summary formatting with various row counts and modification counts
- Test that empty workbooks produce valid debugging artifacts with zero counts

### Property-Based Tests

- Generate random Excel files with varying structures and verify debugging artifacts are always created
- Generate random Excel files and verify output Excel is identical between unfixed and fixed code
- Test that all engine names appear in console output across many normalization runs

### Integration Tests

- Test full pipeline with real Excel file containing person records
- Test that raw_dataset.json can be compared against input Excel to verify extraction
- Test that normalized_dataset.json can be compared against raw_dataset.json to verify engine modifications
- Test that output Excel can be compared against normalized_dataset.json to verify writer correctness
- Test verification workflow: identify defect location using debugging artifacts
