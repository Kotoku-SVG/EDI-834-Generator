# Build a single-file Windows EXE (no Python required for users)

**one EXE** that users can double-click,
with **no Python install** and **no pip installs**.

This is achieved by packaging `mock_834_onefile_app.py` with **PyInstaller** on a Windows machine.

## What you need (builder machine only)
- Windows PC
- Python installed (builder only)
- Internet access (to install PyInstaller once)

## Build steps
1. Put these two files in the same folder on Windows:
   - mock_834_onefile_app.py
   - build_single_exe.bat

2. Double-click `build_single_exe.bat` (or run it in cmd).

3. Your EXE will be here:
   - dist\\mock_834_onefile_app.exe

## Distribute
When packaging is complete, give users:
- dist\\mock_834_onefile_app.exe
- the Excel template (or any workbook they fill out)

Users do NOT need Python.

## Notes / restrictions
The app reads `.xlsx` by parsing OpenXML:
- `.xlsx` only (not .xls, .xlsm)
- best results when cells are typed values (avoid formulas)
