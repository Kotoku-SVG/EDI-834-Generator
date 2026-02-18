@echo off
REM Build a SINGLE-FILE Windows EXE for the Mock 834 app (no Python needed for users).
REM
REM Steps:
REM   1) Put mock_834_onefile_app.py in the same folder as this .bat (or adjust path)
REM   2) Run this .bat
REM
REM Output:
REM   dist\mock_834_onefile_app.exe

setlocal

echo.
echo Installing/Updating PyInstaller...
python -m pip install --upgrade pip
python -m pip install --upgrade pyinstaller

echo.
echo Building single-file EXE (windowed)...
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "mock_834_onefile_app" ^
  "%~dp0mock_834_onefile_app.py"

echo.
echo Build complete.
echo Output: dist\mock_834_onefile_app.exe
echo.
pause
