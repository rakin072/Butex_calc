# -*- coding: utf-8 -*-
"""Resolve application root for source runs and frozen .exe builds."""
from __future__ import annotations

import sys
from pathlib import Path


def app_dir() -> Path:
    """Folder that contains the .exe (frozen) or project files (dev)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent
