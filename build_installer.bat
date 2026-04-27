@echo off
REM ============================================================
REM  build_installer.bat  —  Build EXE then Inno Setup installer
REM
REM  Prerequisites:
REM    pip install pyinstaller
REM    pip install -r requirements.txt
REM    Inno Setup 6 installed (https://jrsoftware.org/isinfo.php)
REM
REM  Output: installer\Output\ExcelNormalization_Setup_1.0.0.exe
REM ============================================================

setlocal

REM ---- Locate iscc.exe -------------------------------------------
REM Try PATH first, then the two standard Inno Setup install locations.
REM We use a separate variable to avoid quoting issues when the path
REM contains spaces (e.g. "C:\Program Files (x86)\...").
set "ISCC="

where iscc >nul 2>&1
if %errorlevel%==0 (
    set "ISCC=iscc"
    goto :found_iscc
)

if exist "C:\Program Files (x86)\Inno Setup 6\iscc.exe" (
    set "ISCC=C:\Program Files (x86)\Inno Setup 6\iscc.exe"
    goto :found_iscc
)

if exist "C:\Program Files\Inno Setup 6\iscc.exe" (
    set "ISCC=C:\Program Files\Inno Setup 6\iscc.exe"
    goto :found_iscc
)

echo ERROR: Inno Setup 6 not found.
echo Download and install from: https://jrsoftware.org/isinfo.php
echo Then re-run this script.
exit /b 1

:found_iscc
echo Using Inno Setup: %ISCC%

REM ----------------------------------------------------------------
echo.
echo ============================================================
echo  Step 1 of 2: Build PyInstaller EXE
echo ============================================================
call build_exe.bat
if errorlevel 1 (
    echo ERROR: EXE build failed. Aborting.
    exit /b 1
)

echo.
echo ============================================================
echo  Step 2 of 2: Compile Inno Setup installer
echo ============================================================
"%ISCC%" installer\ExcelNormalization.iss
if errorlevel 1 (
    echo ERROR: Inno Setup compilation failed.
    exit /b 1
)

echo.
echo ============================================================
echo  Installer ready:
echo  installer\Output\ExcelNormalization_Setup_1.0.0.exe
echo ============================================================
echo.
endlocal
