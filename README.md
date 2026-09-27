# BUTEX Salary Fixation Calculator 2026

Web app for Bangladesh University of Textiles (BUTEX) to look up employees by serial number from an Excel sheet and print the National Pay Scale 2026 fixation form.

## Download & run (end users — no Python)

**Do not use the green “Code → Download ZIP” button** (that is source code only).

1. Open **[Releases](https://github.com/rakin072/Butex_calc/releases/latest)**
2. Download **[BUTEX-Calc.zip](https://github.com/rakin072/Butex_calc/releases/latest/download/BUTEX-Calc.zip)**
3. Unzip the folder (keep everything together)
4. Put your Excel file here: `data\Salary fixation Form.xlsx`
5. Double-click **`BUTEX-Calc.exe`**
6. Browser opens at http://127.0.0.1:8080/
7. Enter serial → select grade → Print / PDF

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
Then upload the new zip to a GitHub Release so end users can download it.

## Project files

| File / folder | Purpose |
|---------------|---------|
| `BUTEX-Calc.exe` (in release zip) | One-click app for users |
| `index.html` | UI + print form |
| `server.py` | Flask backend |
| `excel_data.py` | Live Excel read |
| `data/` | Excel workbook |
| `images/` | BUTEX logo |
| `build_exe.bat` | Rebuild Windows package |

## Notes

- Keep the console window open while using the app.
- If Excel is locked in Excel/OneDrive, save or close it, then try again.
- Employee data may contain personal information — keep the repo private if needed.

## Repository

https://github.com/rakin072/Butex_calc
