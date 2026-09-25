# BUTEX Salary Fixation Calculator 2026

Web app for Bangladesh University of Textiles (BUTEX) to look up employees by serial number from an Excel sheet and print the National Pay Scale 2026 fixation form.

## For end users (easiest — no Python)

1. Download **`release/BUTEX-Calc.zip`** (or the GitHub Release zip).
2. Unzip the folder.
3. Keep Excel here: `data\Salary fixation Form.xlsx`
4. Double-click **`BUTEX-Calc.exe`**
5. Browser opens at http://127.0.0.1:8080/
6. Enter serial → select grade → print.

Read `HOW-TO-RUN.txt` inside the package for details.

## For developers

```bash
pip install -r requirements.txt
python server.py
```

Or double-click `start-server.bat`.

Excel path (only this file):

`data/Salary fixation Form.xlsx`

Optional override:

```bash
set BUTEX_EXCEL=C:\path\to\file.xlsx
```

### Rebuild the .exe

```bash
build_exe.bat
```

Output: `release\BUTEX-Calc\` and `release\BUTEX-Calc.zip`

## Project files

| File / folder | Purpose |
|---------------|---------|
| `BUTEX-Calc.exe` (in release) | One-click app for users |
| `index.html` | UI + print form |
| `server.py` | Flask backend |
| `excel_data.py` | Live Excel read |
| `data/` | Excel workbook |
| `images/` | BUTEX logo |
| `build_exe.bat` | Rebuild Windows package |

## Notes

- Keep the console window open while using the app.
- If Excel is locked in Excel/OneDrive, save/close it and try again.
- Employee data may contain personal information — keep the repo private if needed.

## Repository

https://github.com/rakin072/Butex_calc
