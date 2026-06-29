@echo off
REM Build a standalone Windows .exe of the Athanase Revaluation GUI.
REM Run this from a Windows cmd/PowerShell in the niche_pipeline folder.
REM
REM Prerequisites in this folder before building:
REM   * roiic_dcf.py  -- the valuation engine (copy it here; it is gitignored)
REM   * Python 3 + pip
REM
REM The build bundles: revalue_gui.py and the modules it imports
REM (revalue.py, decompose.py, aip.py, frameworks.py) + roiic_dcf.py + openpyxl.
REM
REM It does NOT bundle your data. After building, place these next to the .exe:
REM   * LTM_baseline.xlsx          (the baseline screener / fundamentals)
REM   * (optional) a scored book   (to value on the researched moat)

setlocal
where python >nul 2>nul || (echo Python not found in PATH & exit /b 1)
if not exist roiic_dcf.py (
  echo ERROR: roiic_dcf.py not found in this folder. Copy the engine here first.
  exit /b 1
)

python -m pip install --upgrade pip
python -m pip install openpyxl pyinstaller

python -m PyInstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name AthanaseRevalue ^
  --add-data "roiic_dcf.py;." ^
  --collect-all openpyxl ^
  revalue_gui.py

echo.
echo Done. Executable is at: dist\AthanaseRevalue.exe
echo Place LTM_baseline.xlsx next to it before handing it to an analyst.
endlocal
