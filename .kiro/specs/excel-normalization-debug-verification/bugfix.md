# Bugfix Requirements Document

## Introduction

The Excel normalization system produces incorrect output in the final Excel file, but the location of the defect is unknown. The defect could be in the extraction layer (Excel → JSON), the normalization engines (NameEngine, GenderEngine, DateEngine, IdentifierEngine), or the Excel writer (JSON → Excel).

Without visibility into intermediate pipeline stages, it is impossible to isolate which component is producing incorrect data. This bugfix adds temporary debugging infrastructure to export datasets at critical pipeline stages and verify that all normalization engines execute correctly.

## Bug Analysis

### Current Behavior (Defect)

1.1 WHEN the normalization pipeline runs THEN the system does not export the raw extracted dataset, making it impossible to verify extraction correctness

1.2 WHEN the normalization pipeline runs THEN the system does not export the normalized dataset before Excel writing, making it impossible to verify engine correctness

1.3 WHEN the normalization pipeline completes THEN the system does not report which engines executed or how many values each engine modified

1.4 WHEN investigating incorrect Excel output THEN the system provides no intermediate artifacts to isolate whether the defect is in extraction, normalization engines, or Excel writer

### Expected Behavior (Correct)

2.1 WHEN the normalization pipeline runs THEN the system SHALL export the raw extracted dataset to `raw_dataset.json` immediately after Excel extraction and before any normalization

2.2 WHEN the normalization pipeline runs THEN the system SHALL execute ALL normalization engines (NameEngine, GenderEngine, DateEngine, IdentifierEngine) on the dataset

2.3 WHEN the normalization pipeline completes THEN the system SHALL export the normalized dataset to `normalized_dataset.json` after all engines execute but before Excel writing

2.4 WHEN the normalization pipeline completes THEN the system SHALL print a console summary showing: number of rows processed, number of columns detected per field, and number of values modified by each engine

2.5 WHEN investigating incorrect Excel output THEN the system SHALL provide intermediate JSON artifacts that allow verification of extraction correctness, engine correctness, and isolation of the defect location

### Unchanged Behavior (Regression Prevention)

3.1 WHEN the normalization pipeline runs with debugging enabled THEN the system SHALL CONTINUE TO produce the same normalized Excel output file

3.2 WHEN the normalization pipeline runs THEN the system SHALL CONTINUE TO execute all engines in the correct order: NameEngine, GenderEngine, DateEngine, IdentifierEngine

3.3 WHEN the normalization pipeline runs THEN the system SHALL CONTINUE TO preserve original values and create corrected fields with the "_corrected" suffix

3.4 WHEN the normalization pipeline runs THEN the system SHALL CONTINUE TO log normalization statistics and errors to the log file


## Bug Condition Analysis

### Bug Condition Function

The bug condition identifies when the pipeline lacks debugging visibility:

```pascal
FUNCTION isBugCondition(pipeline_execution)
  INPUT: pipeline_execution of type PipelineExecution
  OUTPUT: boolean
  
  // Returns true when debugging artifacts are missing
  RETURN (NOT pipeline_execution.exports_raw_json) OR
         (NOT pipeline_execution.exports_normalized_json) OR
         (NOT pipeline_execution.reports_engine_statistics)
END FUNCTION
```

### Property Specification - Fix Checking

After the fix, the pipeline must provide debugging visibility:

```pascal
// Property: Fix Checking - Debugging Artifacts Present
FOR ALL pipeline_execution WHERE isBugCondition(pipeline_execution) DO
  result ← execute_pipeline_with_debug'(pipeline_execution)
  
  ASSERT file_exists("raw_dataset.json") AND
         file_exists("normalized_dataset.json") AND
         console_output_contains_engine_statistics(result) AND
         raw_json_represents_exact_extraction(result) AND
         normalized_json_contains_all_corrected_fields(result)
END FOR
```

### Property Specification - Preservation Checking

For pipeline executions that already work correctly, behavior must be unchanged:

```pascal
// Property: Preservation Checking - Pipeline Output Unchanged
FOR ALL pipeline_execution WHERE NOT isBugCondition(pipeline_execution) DO
  original_output ← execute_pipeline(pipeline_execution)
  debug_output ← execute_pipeline_with_debug'(pipeline_execution)
  
  ASSERT original_output.excel_file = debug_output.excel_file AND
         original_output.normalized_data = debug_output.normalized_data AND
         original_output.engine_execution_order = debug_output.engine_execution_order
END FOR
```

### Verification Strategy

The debugging infrastructure enables systematic verification:

1. **Extraction Verification**: Compare `raw_dataset.json` against Excel file to verify extraction correctness
2. **Engine Verification**: Compare `raw_dataset.json` vs `normalized_dataset.json` to verify each engine modified the correct fields
3. **Writer Verification**: Compare `normalized_dataset.json` vs output Excel file to verify writer correctness
4. **Defect Isolation**: Identify which component (extraction, engines, or writer) produces incorrect data

### Engine Execution Verification

All engines must execute and report modifications:

```pascal
FUNCTION verify_all_engines_executed(console_output)
  INPUT: console_output of type String
  OUTPUT: boolean
  
  required_engines ← ["NameEngine", "GenderEngine", "DateEngine", "IdentifierEngine"]
  
  FOR EACH engine IN required_engines DO
    IF NOT console_output.contains(engine + " modified:") THEN
      RETURN false
    END IF
  END FOR
  
  RETURN true
END FUNCTION
```
