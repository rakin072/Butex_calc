@echo off
cd /d "%~dp0"
echo Building BUTEX-Calc.exe ...
python -m pip install -r requirements.txt pyinstaller -q
if errorlevel 1 (
  echo Failed to install build tools.
  exit /b 1
)

python -m PyInstaller --noconfirm --clean --onefile --console --name BUTEX-Calc ^
  --hidden-import=excel_data ^
  --hidden-import=app_paths ^
  --hidden-import=openpyxl ^
  --hidden-import=unicodeconverter ^
  --collect-all openpyxl ^
  --collect-all unicodeconverter ^
  --collect-all flask ^
  server.py

if errorlevel 1 (
  echo Build failed.
  exit /b 1
)

echo Assembling release package...
set PKG=release\BUTEX-Calc
if exist "%PKG%" rmdir /s /q "%PKG%"
mkdir "%PKG%"
mkdir "%PKG%\data"
mkdir "%PKG%\images"
mkdir "%PKG%\vendor"

copy /Y "dist\BUTEX-Calc.exe" "%PKG%\BUTEX-Calc.exe" >nul
copy /Y "index.html" "%PKG%\index.html" >nul
copy /Y "images\*.*" "%PKG%\images\" >nul
copy /Y "vendor\html2pdf.bundle.min.js" "%PKG%\vendor\" >nul
if exist "data\Salary fixation Form.xlsx" (
  copy /Y "data\Salary fixation Form.xlsx" "%PKG%\data\" >nul
) else if exist "data\Salary fixation Form (2).xlsx" (
  copy /Y "data\Salary fixation Form (2).xlsx" "%PKG%\data\Salary fixation Form.xlsx" >nul
)
copy /Y "HOW-TO-RUN.txt" "%PKG%\HOW-TO-RUN.txt" >nul

echo.
echo Done: %PKG%\BUTEX-Calc.exe
echo Zip the folder "%PKG%" and share it.
