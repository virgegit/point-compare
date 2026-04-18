"""Validation functions for point-compare."""

from pathlib import Path
import pandas as pd


def validate_workbook(path_str: str) -> tuple:
    """Validate that path points to a valid Excel file."""
    p = Path(path_str)
    if not p.exists():
        return False, f"File not found: {p}"
    if p.suffix.lower() not in (".xlsx", ".xlsm"):
        return False, "Only .xlsx and .xlsm Excel files are supported"
    return True, ""


def list_sheets(path: str) -> list:
    """Return worksheet names from an Excel file."""
    try:
        xl = pd.ExcelFile(path)
        return xl.sheet_names
    except Exception:
        return []
