# BUTEX Salary Fixation Calculator 2026

Web app for Bangladesh University of Textiles (BUTEX) to look up employees by serial number from an Excel sheet and print the National Pay Scale 2026 fixation form.

## Requirements

- Windows PC
- [Python 3.10+](https://www.python.org/downloads/) (check **Add Python to PATH** during install)
- Excel file with sheet name `Inforation` (same structure as the salary fixation workbook)

## Quick start (for users)

1. Unzip / clone this folder.
2. Place your Excel file at:

   `data/Salary fixation Form.xlsx`

3. Double-click `start-server.bat`  
   (first run installs packages automatically).
4. Browser opens at [http://127.0.0.1:8080/](http://127.0.0.1:8080/)
5. Enter **serial number** → **তথ্য আনুন** → select **প্রাপ্য গ্রেড** → **প্রিন্ট / PDF**.

To stop: close the black terminal window.

## Manual start

```bash
pip install -r requirements.txt
python server.py
```

Optional: set a custom Excel path:

```bash
set BUTEX_EXCEL=C:\path\to\your\file.xlsx
python server.py
```

## Project files

| File | Purpose |
|------|---------|
| `index.html` | UI + print form |
| `server.py` | Flask backend (live Excel lookup) |
| `excel_data.py` | Excel read + Bijoy→Unicode |
| `data/` | Put `Salary fixation Form.xlsx` here |
| `start-server.bat` | One-click start for Windows |
| `employees.js` / `employees.json` | Optional static snapshot (not required for live mode) |

## Notes

- Do not open `index.html` directly as a file; use the server URL.
- If Excel is open and locked, save/close it and try again.
- Employee data may contain personal information (NID, phone). Keep the repo private if needed.

## Repository

https://github.com/rakin072/Butex_calc
