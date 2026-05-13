# Packaging Rules

## Purpose

Document how the project is packaged for local execution and installer use.

## Scope

- `pyproject.toml`
- `ExcelNormalization.spec`
- `build_exe.bat`
- `build_installer.bat`
- `installer/Excelstandardization.iss`

## Main Files

- `pyproject.toml`
- `ExcelNormalization.spec`
- `build_exe.bat`
- `build_installer.bat`
- `installer/Excelstandardization.iss`

## Responsibilities

- Describe how the app is bundled for distribution.
- Ensure required dependencies are included in the packaged runtime.
- Keep the installer aligned with the app's actual runtime folders.

## Data Flow

1. Dependencies are declared in `pyproject.toml`.
2. PyInstaller bundles the app.
3. Inno Setup packages the bundle.
4. The installer creates the runtime folders.

## Contracts

- The packaged app must still support the active web/session flow.
- Required Excel reader/writer dependencies must be present in the bundle.

## What Must Never Change

- Packaging must not alter standardization business rules.
- The installer must not overwrite source workbooks.

## Current Behavior

- The project uses a PyInstaller spec and an Inno Setup script.
- The build scripts clean previous artifacts and build the EXE plus installer.

## Known Limitations

- Packaging behavior is tied to the project's current dependency set.
- The runtime may still be sensitive to missing binaries or library versions.

## Tests That Should Cover It

- build smoke checks
- installed/exe startup checks
- path creation checks for output/upload/work folders

## Open Questions / Future Improvements

- Whether to add a formal packaging smoke-test script.
- Whether to document runtime folder creation in one shared place for users.
