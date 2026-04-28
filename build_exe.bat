@echo off
REM ============================================================
REM  build_exe.bat  —  Build the Excelstandardization.exe bundle
REM
REM  Prerequisites (run once):
REM    pip install pyinstaller
REM    pip install -r requirements.txt
REM
REM  Output: dist\Excelstandardization\Excelstandardization.exe
REM          dist\Excelstandardization\_internal\webapp\static\
REM          dist\Excelstandardization\_internal\webapp\templates\
REM ============================================================

setlocal

echo.
echo ============================================================
echo  Excel standardization — EXE Build
echo ============================================================
echo.

REM Clean previous build artefacts
if exist build\Excelstandardization (
    echo Cleaning previous build...
    rmdir /s /q build\Excelstandardization
)
if exist dist\Excelstandardization (
    echo Cleaning previous dist...
    rmdir /s /q dist\Excelstandardization
)

REM Remove stale __pycache__ so PyInstaller sees fresh bytecode
echo Removing stale bytecode...
for /d /r . %%d in (__pycache__) do (
    if exist "%%d" rmdir /s /q "%%d"
)

echo.
echo Running PyInstaller...
pyinstaller Excelstandardization.spec --noconfirm

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

if not exist "dist\Excelstandardization\Excelstandardization.exe" (
    echo ERROR: Excelstandardization.exe not found in dist folder.
    exit /b 1
)
if not exist "dist\Excelstandardization\_internal\webapp\static\app.js" (
    echo ERROR: _internal\webapp\static\app.js missing from bundle.
    exit /b 1
)
if not exist "dist\Excelstandardization\_internal\webapp\templates\index.html" (
    echo ERROR: _internal\webapp\templates\index.html missing from bundle.
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete.
echo  Executable: dist\Excelstandardization\Excelstandardization.exe
echo ============================================================
echo.
endlocal
