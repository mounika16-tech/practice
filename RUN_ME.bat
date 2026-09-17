@echo off
REM ============================================================
REM  Excel Updater - double-click launcher (Windows)
REM  Put your .xlsx file in this same folder, then double-click.
REM ============================================================
setlocal

where python >nul 2>&1
if errorlevel 1 (
  echo.
  echo ERROR: Python was not found on this machine.
  echo Ask IT to install Python 3 ^(python.org^), or run this on a machine that has it.
  echo.
  pause
  exit /b 1
)

echo.
set /p WB="Excel file name (e.g. MyWorkbook.xlsx): "
if "%WB%"=="" ( echo No file given. & pause & exit /b 1 )
if not exist "%WB%" ( echo File not found: %WB% & pause & exit /b 1 )

echo.
set /p UID="Unique id (e.g. BRP1E): "
if "%UID%"=="" ( echo No unique id given. & pause & exit /b 1 )

echo.
echo === STEP 1: PREVIEW for %UID% (nothing is saved) ===
python excel_updater.py --workbook "%WB%" --uid "%UID%" --config template_config.json --dry-run
if errorlevel 1 ( pause & exit /b 1 )

echo.
set /p GO="Look OK? Type Y to write the output file: "
if /i not "%GO%"=="Y" ( echo Cancelled. Nothing was changed. & pause & exit /b 0 )

echo.
echo === STEP 2: WRITING OUTPUT for %UID% ===
python excel_updater.py --workbook "%WB%" --uid "%UID%" --config template_config.json --output "UPDATED_%WB%"
echo.
echo Done. Your original file was not modified.
pause
