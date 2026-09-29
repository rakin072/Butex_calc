# -*- coding: utf-8 -*-
"""Live Excel backend for BUTEX salary fixation calculator."""
from __future__ import annotations

import threading
import time
import webbrowser

from flask import Flask, jsonify, send_from_directory

from app_paths import app_dir
from excel_data import excel_stats, find_employee, resolve_excel_path

BASE_DIR = app_dir()
app = Flask(__name__, static_folder=str(BASE_DIR), static_url_path="")
PORT = 8080
URL = f"http://127.0.0.1:{PORT}/"


@app.get("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.get("/api/health")
def health():
    try:
        path = resolve_excel_path()
        return jsonify({"ok": True, "excelPath": str(path)})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.get("/api/stats")
def stats():
    try:
        return jsonify(excel_stats())
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except PermissionError as exc:
        return jsonify({"error": str(exc)}), 423
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.get("/api/employee/<int:serial>")
def employee(serial: int):
    try:
        emp = find_employee(serial)
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except PermissionError as exc:
        return jsonify({"error": str(exc)}), 423
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

    if not emp:
        return jsonify({"error": f"ক্রমিক নং {serial} পাওয়া যায়নি।"}), 404
    return jsonify(emp)


def _open_browser():
    time.sleep(1.2)
    webbrowser.open(URL)


if __name__ == "__main__":
    print("=" * 50)
    print(" BUTEX Salary Fixation Calculator")
    print("=" * 50)
    print("App folder:", BASE_DIR)
    try:
        print("Excel:", resolve_excel_path())
    except Exception as exc:
        print("Excel warning:", exc)
        print("Put the file at:", BASE_DIR / "data" / "Salary fixation Form.xlsx")
        print("Or:", BASE_DIR / "data" / "Salary fixation Form (2).xlsx")
    print("Opening", URL)
    print("Keep this window open while using the app.")
    print("Close this window to stop the server.")
    print("=" * 50)
    threading.Thread(target=_open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False)
