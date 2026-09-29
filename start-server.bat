@echo off
cd /d "%~dp0"
echo ========================================
echo  BUTEX Salary Fixation Calculator
echo ========================================
echo.

where python >nul 2>&1
if errorlevel 1 (
  echo ERROR: Python not found. Install Python and tick "Add to PATH".
  echo Or use the release package BUTEX-Calc.exe instead.
  echo https://www.python.org/downloads/
  pause
  exit /b 1
)

echo Installing / updating packages...
python -m pip install -r requirements.txt -q
if errorlevel 1 (
  echo Package install failed.
  pause
  exit /b 1
)

if not exist "data\Salary fixation Form.xlsx" if not exist "data\Salary fixation Form (2).xlsx" (
  echo.
  echo WARNING: Excel file not found in data\
  echo Put: "Salary fixation Form.xlsx" or "Salary fixation Form (2).xlsx"
  echo.
)

echo Starting server at http://127.0.0.1:8080/
echo.
python server.py
pause
