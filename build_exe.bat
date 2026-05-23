@echo off
REM Build a standalone Windows .exe of the SQL explorer GUI.
REM Run this from a Windows cmd or PowerShell in the project folder.
REM Requires: Python 3, pip, and the ODBC Driver 18 for SQL Server installed
REM on the build machine.

setlocal
where python >nul 2>nul || (echo Python not found in PATH & exit /b 1)

python -m pip install --upgrade pip
python -m pip install pyodbc pyinstaller

python -m PyInstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name AthanaseSQLExplorer ^
  app_gui.py

echo.
echo Done. Executable is at: dist\AthanaseSQLExplorer.exe
endlocal
