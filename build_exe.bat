@echo off
REM ============================================================
REM  build_exe.bat  —  Build the ExcelNormalization.exe bundle
REM
REM  Prerequisites (run once):
REM    pip install pyinstaller
REM    pip install -r requirements.txt
REM
REM  Output: dist\ExcelNormalization\ExcelNormalization.exe
REM          dist\ExcelNormalization\_internal\webapp\static\
REM          dist\ExcelNormalization\_internal\webapp\templates\
REM ============================================================

setlocal

echo.
echo ============================================================
echo  Excel Normalization — EXE Build
echo ============================================================
echo.

REM Clean previous build artefacts
if exist build\ExcelNormalization (
    echo Cleaning previous build...
    rmdir /s /q build\ExcelNormalization
)
if exist dist\ExcelNormalization (
    echo Cleaning previous dist...
    rmdir /s /q dist\ExcelNormalization
)

REM Remove stale __pycache__ so PyInstaller sees fresh bytecode
echo Removing stale bytecode...
for /d /r . %%d in (__pycache__) do (
    if exist "%%d" rmdir /s /q "%%d"
)

echo.
echo Running PyInstaller...
pyinstaller ExcelNormalization.spec --noconfirm

if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed.
    exit /b 1
)

REM ---------------------------------------------------------------
REM Smoke-test: PyInstaller 6+ places support files under _internal
REM ---------------------------------------------------------------
echo.
echo Verifying output...

if not exist "dist\ExcelNormalization\ExcelNormalization.exe" (
    echo ERROR: ExcelNormalization.exe not found in dist folder.
    exit /b 1
)
if not exist "dist\ExcelNormalization\_internal\webapp\static\app.js" (
    echo ERROR: _internal\webapp\static\app.js missing from bundle.
    exit /b 1
)
if not exist "dist\ExcelNormalization\_internal\webapp\templates\index.html" (
    echo ERROR: _internal\webapp\templates\index.html missing from bundle.
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete.
echo  Executable: dist\ExcelNormalization\ExcelNormalization.exe
echo ============================================================
echo.
endlocal
