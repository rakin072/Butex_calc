@echo off
cd /d "%~dp0"
echo ========================================
echo  BUTEX Salary Fixation Calculator
echo ========================================
echo.

where python >nul 2>&1
if errorlevel 1 (
  echo ERROR: Python not found. Install Python and tick "Add to PATH".
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

if not exist "data\Salary fixation Form.xlsx" (
  echo.
  echo WARNING: data\Salary fixation Form.xlsx not found.
  echo Put your Excel file there, then run this again.
  echo.
)

echo Starting server...
echo Open http://127.0.0.1:8080/ in your browser.
echo.
start "" "http://127.0.0.1:8080/"
python server.py
pause
