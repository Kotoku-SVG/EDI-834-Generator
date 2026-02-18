@echo off
REM --- Mock 834 Generator Launcher ---
REM Put this .bat in the SAME folder as:
REM   - mock_834_generator.py
REM   - mock_834_generator_template.xlsx (or your filled-in workbook)
REM Double-click this .bat to run and keep the window open.

set "IN_FILE=mock_834_generator_template.xlsx"
set "OUT_FILE=out_834.txt"

echo.
echo Running Mock 834 Generator...
echo   Input : %IN_FILE%
echo   Output: %OUT_FILE%
echo.

python mock_834_generator.py --in "%IN_FILE%" --out "%OUT_FILE%"

echo.
if %ERRORLEVEL% EQU 0 (
  echo ✅ Done! Created %OUT_FILE%
) else (
  echo ❌ There was an error running the generator. See messages above.
)

echo.
pause
