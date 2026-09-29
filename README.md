# BUTEX Salary Fixation Calculator 2026

Web app for Bangladesh University of Textiles (BUTEX) to look up employees by serial number from an Excel sheet and print the National Pay Scale 2026 fixation form.

## Download & run (end users — no Python)

1. On this page, click the green **Code** button → **Download ZIP**
2. Unzip `Butex_calc-main.zip`
3. Open the folder: **`release\BUTEX-Calc`**
4. Double-click **`BUTEX-Calc.exe`**
5. Browser opens at http://127.0.0.1:8080/
6. Enter serial → select grade → Print / PDF

Keep Excel here (inside that folder): `data\Salary fixation Form.xlsx`

Read `HOW-TO-RUN.txt` inside the package for details.

## For developers

```bash
pip install -r requirements.txt
python server.py
```

Or double-click `start-server.bat`.

Excel path (first match wins):

- `data/Salary fixation Form.xlsx`
- `data/Salary fixation Form (2).xlsx`

Optional override:

```bash
set BUTEX_EXCEL=C:\path\to\file.xlsx
```

Print uses the official **Fixation** sheet layout. Use **PDF ডাউনলোড** to save the form, or **প্রিন্ট** for the browser print dialog.

### Rebuild the .exe

```bash
build_exe.bat
```

Output: `release\BUTEX-Calc\` (committed for end users) and `release\BUTEX-Calc.zip`  
After rebuilding, commit the updated `release\BUTEX-Calc\` folder and push.

## Project files

| File / folder | Purpose |
|---------------|---------|
| `release/BUTEX-Calc/BUTEX-Calc.exe` | One-click app for users |
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
