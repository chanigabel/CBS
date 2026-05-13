# Agent: Packaging

## Mission

Review and maintain the packaged runtime, EXE build, and installer flow.

## Files To Inspect First

- `docs/project_rules/PACKAGING_RULES.md`
- `pyproject.toml`
- `ExcelNormalization.spec`
- `build_exe.bat`
- `build_installer.bat`
- `installer/Excelstandardization.iss`

## Rules To Follow

- Keep packaging aligned with the active web/session runtime.
- Keep dependencies explicit.

## What The Agent May Change

- Packaging config, build scripts, and docs when requested.

## What The Agent Must Not Change

- Runtime business behavior.
- Excel input/output contracts.

## Required Tests

- packaging smoke tests if available
- startup verification for bundled runtime
- path creation checks

## Regression Checklist

- bundled app still starts
- installer still creates runtime folders
- dependency list still covers Excel readers/writers

## Expected Output Format

- packaging findings
- dependency findings
- smoke-test results

## Safety Constraints

- Do not break the dev/runtime path while adjusting packaging.
- Do not assume the bundle contains unlisted dependencies.

