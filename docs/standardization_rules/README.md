# Standardization Rules

This folder documents the active standardization behavior in the current codebase.
It is descriptive, not speculative: rules are based on the implemented engines,
processing pipeline, export services, UI behavior, existing tests, and the root
`DATE_RULES.md` document.

## Rule Status Labels

- **Approved rule**: behavior that is explicitly implemented, covered by tests, or already documented as a rule.
- **Current behavior**: behavior observed in code or tests that may be intentional but is not separately approved as a business rule.
- **Needs approval**: unclear behavior, missing source data, or a code/documentation conflict that requires product or business approval before being treated as a rule.
- **Potential issue**: behavior that may surprise users or may conflict with a desired future contract.

## Documents

- `DATE_RULES_REFERENCE.md` points to the existing root `DATE_RULES.md`.
- `NAME_RULES.md` covers name cleanup and first/father name last-name removal.
- `GENDER_RULES.md` covers gender normalization and gender status behavior.
- `IDENTIFIER_RULES.md` covers Israeli ID and passport normalization.
- `TEXT_CLEANUP_RULES.md` covers the generic text/name cleanup processor.
- `INSTITUTION_RULES.md` covers institution-report validation and Mosad-related behavior.
- `EXPORT_RULES.md` covers export schema, row filtering, field mapping, and web export behavior.
- `PIPELINE_RULES.md` covers orchestration, corrected-field creation, UI/grid visibility, and API flow.

## Source Of Truth

Use these files with the implementation, not instead of it. When behavior is disputed,
inspect the current code first:

- `src/excel_standardization/engines/`
- `src/excel_standardization/processing/`
- `src/excel_standardization/export/export_engine.py`
- `src/excel_standardization/validation/institution_report_validator.py`
- `webapp/services/`
- `webapp/api/`
- `tests/`

## Final Principles

Original values are immutable. Standardization writes corrected values into
`*_corrected` fields and writes user-visible explanations into status fields.
Export must prefer corrected fields where the active schema maps to them.
Ambiguous or unsupported behavior must be documented as current behavior or
needs approval, not silently converted into a new rule.
