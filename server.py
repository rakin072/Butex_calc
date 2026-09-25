# -*- coding: utf-8 -*-
"""Live Excel backend for BUTEX salary fixation calculator."""
from __future__ import annotations

from pathlib import Path

from flask import Flask, jsonify, send_from_directory

from excel_data import excel_stats, find_employee, resolve_excel_path

BASE_DIR = Path(__file__).resolve().parent
app = Flask(__name__, static_folder=str(BASE_DIR), static_url_path="")


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


if __name__ == "__main__":
    print("BUTEX live Excel server")
    try:
        print("Excel:", resolve_excel_path())
    except Exception as exc:
        print("Excel warning:", exc)
    print("Open http://127.0.0.1:8080/")
    app.run(host="127.0.0.1", port=8080, debug=False)
