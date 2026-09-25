# -*- coding: utf-8 -*-
"""Shared helpers: read live employee rows from the Salary Fixation Excel file."""
from __future__ import annotations

import os
import re
from pathlib import Path

import openpyxl
import unicodeconverter as uc

BASE_DIR = Path(__file__).resolve().parent

# Prefer project data/ (for sharing), then Downloads. Override with BUTEX_EXCEL.
DEFAULT_CANDIDATES = [
    BASE_DIR / "data" / "Salary fixation Form.xlsx",
    Path(r"c:\Users\omarr\Downloads\Salary fixation Form (1).xlsx"),
]

SCALES_OLD = {
    1: [78000],
    2: [66000, 68480, 71050, 73720, 76490],
    3: [56500, 58760, 61120, 63570, 66120, 68770, 71530, 74400],
    4: [50000, 52000, 54080, 56250, 58500, 60840, 63280, 65820, 68460, 71200],
    5: [43000, 44940, 46970, 49090, 51300, 53650, 56030, 58560, 61200, 63960, 66840, 69850],
    6: [35500, 37280, 39150, 41110, 43170, 45330, 47600, 49980, 52480, 55110, 57870, 60770, 63810, 67010],
    7: [29000, 30450, 31980, 33580, 35260, 37030, 38890, 40840, 42890, 45040, 47300, 49670, 52160, 54770, 57510, 60390, 63410],
    8: [23000, 24150, 25360, 26630, 27970, 29370, 30840, 32390, 34010, 35720, 37510, 39390, 41360, 43430, 45610, 47900, 50300, 52820, 55470],
    9: [22000, 23100, 24260, 25480, 26760, 28100, 29510, 30990, 32540, 34170, 35880, 37680, 39570, 41550, 43630, 45810, 48120, 50530, 53060],
    10: [16000, 16800, 17640, 18530, 19460, 20440, 21470, 22550, 23680, 24870, 26120, 27430, 28810, 30260, 31780, 33370, 35040, 36800, 38640],
    11: [12500, 13130, 13790, 14480, 15210, 15980, 16780, 17620, 18510, 19440, 20420, 21440, 22530, 23660, 24850, 26100, 27410, 28790, 30230],
    12: [11300, 11870, 12470, 13100, 13760, 14450, 15180, 15940, 16740, 17580, 18460, 19390, 20360, 21380, 22450, 23580, 24760, 26000, 27300],
    13: [11000, 11550, 12130, 12740, 13380, 14050, 14760, 15500, 16280, 17100, 17960, 18860, 19810, 20810, 21860, 22960, 24110, 25320, 26590],
    14: [10200, 10710, 11250, 11820, 12420, 13050, 13710, 14400, 15120, 15880, 16680, 17520, 18400, 19320, 20290, 21310, 22380, 23500, 24680],
    15: [9700, 10190, 10700, 11240, 11810, 12410, 13040, 13700, 14390, 15110, 15870, 16670, 17510, 18390, 19310, 20280, 21300, 22370, 23490],
    16: [9300, 9770, 10260, 10780, 11320, 11890, 12490, 13120, 13780, 14470, 15200, 15960, 16760, 17600, 18480, 19410, 20390, 21410, 22490],
    17: [9000, 9450, 9930, 10430, 10960, 11510, 12090, 12700, 13380, 14050, 14720, 15480, 16280, 17060, 17920, 18820, 19770, 20760, 21800],
    18: [8800, 9240, 9710, 10200, 10710, 11250, 11820, 12420, 13050, 13710, 14400, 15120, 15880, 16680, 17520, 18400, 19320, 20290, 21310],
    19: [8500, 8930, 9380, 9850, 10350, 10870, 11420, 12000, 12600, 13230, 13890, 14600, 15330, 16100, 16910, 17760, 18650, 19590, 20570],
    20: [8250, 8670, 9110, 9570, 10050, 10560, 11090, 11650, 12240, 12860, 13510, 14190, 14900, 15650, 16440, 17270, 18140, 19050, 20010],
}

# Known Excel typos → nearest scale step
SALARY_ALIASES = {53610: 53650}

SALARY_LOOKUP = {}
for grade in range(20, 0, -1):
    for idx, pay in enumerate(SCALES_OLD[grade]):
        SALARY_LOOKUP[pay] = (grade, idx)


def resolve_excel_path() -> Path:
    env = os.environ.get("BUTEX_EXCEL", "").strip()
    if env:
        path = Path(env)
        if path.is_file():
            return path
        raise FileNotFoundError(f"BUTEX_EXCEL not found: {path}")
    for path in DEFAULT_CANDIDATES:
        if path.is_file():
            return path
    raise FileNotFoundError(
        "Excel file not found. Place it at data/Salary fixation Form.xlsx "
        "or set BUTEX_EXCEL to the full path."
    )


def to_unicode(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    # Already Unicode Bengali?
    if any("\u0980" <= ch <= "\u09FF" for ch in text):
        return re.sub(r"\s+", " ", text).strip()
    try:
        converted = uc.convert_bijoy_to_unicode(text)
    except Exception:
        converted = text
    converted = re.sub(r"\s+", " ", converted).strip()
    converted = converted.replace("দদপ্তর", "দপ্তর")
    if converted.endswith("প্তর") and not converted.endswith("দপ্তর"):
        converted = converted[:-3] + "দপ্তর"
    converted = converted.replace("ো", "ো")
    converted = converted.replace("মোঃ", "মোঃ").replace("মো:", "মোঃ")
    return converted


def _phone_str(phone) -> str:
    if phone is None:
        return ""
    if isinstance(phone, float) and phone == int(phone):
        return str(int(phone))
    return str(phone).strip()


def _basic_pay(value):
    if value is None or value == "":
        return None
    try:
        pay = int(value)
    except (TypeError, ValueError):
        return None
    return SALARY_ALIASES.get(pay, pay)


def _row_to_employee(serial: int, row_vals) -> dict:
    name, post, department, phone, nid, basic_raw = row_vals
    basic = _basic_pay(basic_raw)
    grade = step = None
    if basic in SALARY_LOOKUP:
        grade, step = SALARY_LOOKUP[basic]
    return {
        "serial": serial,
        "name": to_unicode(name),
        "post": to_unicode(post),
        "department": to_unicode(department),
        "phone": _phone_str(phone),
        "nid": "" if nid is None else str(nid).strip(),
        "basic": basic,
        "grade": grade,
        "step": step,
    }


def find_employee(serial: int) -> dict | None:
    """Read Excel on each call and return one employee by serial (ক্রমিক নং)."""
    path = resolve_excel_path()
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except PermissionError as exc:
        raise PermissionError(
            "Excel file is locked (close it in Excel/OneDrive and try again)."
        ) from exc

    try:
        if "Inforation" not in wb.sheetnames:
            raise ValueError("Sheet 'Inforation' not found in Excel file.")
        ws = wb["Inforation"]
        for row in ws.iter_rows(min_row=2, max_col=7, values_only=True):
            raw_serial = row[0]
            if raw_serial is None:
                continue
            try:
                row_serial = int(raw_serial)
            except (TypeError, ValueError):
                continue
            if row_serial == serial:
                return _row_to_employee(row_serial, row[1:7])
    finally:
        wb.close()
    return None


def excel_stats() -> dict:
    path = resolve_excel_path()
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except PermissionError as exc:
        raise PermissionError(
            "Excel file is locked (close it in Excel/OneDrive and try again)."
        ) from exc

    count = 0
    max_serial = 0
    try:
        ws = wb["Inforation"]
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            raw = row[0]
            if raw is None:
                continue
            try:
                serial = int(raw)
            except (TypeError, ValueError):
                continue
            count += 1
            if serial > max_serial:
                max_serial = serial
    finally:
        wb.close()

    return {
        "count": count,
        "maxSerial": max_serial,
        "excelPath": str(path),
        "live": True,
    }
