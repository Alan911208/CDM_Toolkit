"""
CDM Excel Toolkit Plus — Python supplement adding features discovered
during VBA analysis that are missing from cdm_toolkit.py.

New features:
  1. File directory listing → Excel with hyperlinks
  2. Folder tree → Excel with indentation
  3. Track changes with real Excel comments (via COM/xlwings)
  4. Export .bas module into .xlsm workbook
"""

from __future__ import annotations
import os
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.worksheet.hyperlink import Hyperlink


# ============================================================
#  Style constants (match cdm_toolkit.py)
# ============================================================
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TOC_FONT = Font(color="0563C1", underline="single", size=11)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
GREEN_FILL = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
RED_BORDER = Border(
    left=Side(style="medium", color="FF0000"),
    right=Side(style="medium", color="FF0000"),
    top=Side(style="medium", color="FF0000"),
    bottom=Side(style="medium", color="FF0000"),
)


# ============================================================
#  1. File Directory Listing → Excel
# ============================================================
def generate_file_listing(
    folder_path: str,
    output_path: str,
    include_subfolders: bool = True,
    file_pattern: str = "*.*",
) -> Workbook:
    """
    Generate an Excel file listing all files in a folder with hyperlinks.

    Parameters
    ----------
    folder_path : str
        Root folder to scan.
    output_path : str
        Where to save the Excel file.
    include_subfolders : bool
        Recursively scan subfolders (default True).
    file_pattern : str
        File filter, e.g. "*.xlsx" or "*.pdf" (default "*.*").

    Returns
    -------
    Workbook
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "File Listing"

    # Header
    ws.cell(row=1, column=1, value="Folder Path").font = HEADER_FONT
    ws.cell(row=1, column=1).fill = HEADER_FILL
    ws.cell(row=1, column=2, value="File Name").font = HEADER_FONT
    ws.cell(row=1, column=2).fill = HEADER_FILL

    # Parse file pattern
    if file_pattern in ("*.*", "*"):
        ext_filter = None
    elif file_pattern.startswith("*."):
        ext_filter = file_pattern[1:].lower()  # ".xlsx"
    else:
        ext_filter = file_pattern.lower()

    row_idx = 2

    def _scan(path: str) -> None:
        nonlocal row_idx
        try:
            for entry in os.scandir(path):
                if entry.is_file():
                    name = entry.name
                    if ext_filter is None or name.lower().endswith(ext_filter):
                        ws.cell(row=row_idx, column=1, value=os.path.dirname(entry.path))
                        cell = ws.cell(row=row_idx, column=2, value=name)
                        cell.font = TOC_FONT
                        cell.hyperlink = Hyperlink(
                            ref=cell.coordinate,
                            target=entry.path,
                            display=name,
                        )
                        row_idx += 1
                elif entry.is_dir() and include_subfolders:
                    _scan(entry.path)
        except PermissionError:
            pass  # Skip inaccessible folders

    _scan(folder_path)

    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 50
    wb.save(output_path)
    return wb


# ============================================================
#  2. Folder Tree → Excel
# ============================================================
def generate_folder_tree(
    folder_path: str,
    output_path: str,
) -> Workbook:
    """
    Generate a recursive tree-view folder structure as Excel.

    Parameters
    ----------
    folder_path : str
        Root folder to scan.
    output_path : str
        Where to save the Excel file.

    Returns
    -------
    Workbook
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Folder Tree"

    ws.cell(row=1, column=1, value="Folder Tree Structure").font = Font(bold=True, size=14)
    ws.cell(row=2, column=1, value=os.path.abspath(folder_path))
    ws.cell(row=2, column=1).fill = HEADER_FILL
    ws.cell(row=2, column=1).font = HEADER_FONT

    row_idx = 3

    def _build_tree(path: str, depth: int = 0) -> None:
        nonlocal row_idx
        indent = "    " * depth
        try:
            entries = sorted(os.scandir(path), key=lambda e: (not e.is_dir(), e.name.lower()))
            for entry in entries:
                if entry.is_dir():
                    ws.cell(row=row_idx, column=1, value=f"{indent}|- {entry.name}\\")
                    ws.cell(row=row_idx, column=1).fill = YELLOW_FILL
                    row_idx += 1
                    _build_tree(entry.path, depth + 1)
                else:
                    cell = ws.cell(row=row_idx, column=1, value=f"{indent}|  {entry.name}")
                    cell.font = TOC_FONT
                    cell.hyperlink = Hyperlink(
                        ref=cell.coordinate,
                        target=entry.path,
                        display=entry.name,
                    )
                    row_idx += 1
        except PermissionError:
            ws.cell(row=row_idx, column=1, value=f"{indent}[Access Denied]")
            row_idx += 1

    _build_tree(folder_path)
    ws.column_dimensions["A"].width = 80
    wb.save(output_path)
    return wb


# ============================================================
#  3. Track Changes with Excel Comments
# ============================================================
def track_changes_with_comments(
    wb: Workbook,
    original_path: str,
) -> int:
    """
    Track changes by comparing a workbook against an original copy.
    Uses openpyxl color marking (COM/xlwings not required for basic use).

    Parameters
    ----------
    wb : Workbook (openpyxl)
        Current workbook (may be unsaved).
    original_path : str
        Path to the original workbook for comparison.

    Returns
    -------
    int — number of changed cells detected.
    """
    orig_wb = openpyxl.load_workbook(original_path, data_only=True)
    changes = 0

    for ws in wb.worksheets:
        if ws.title not in orig_wb.sheetnames:
            continue
        orig_ws = orig_wb[ws.title]

        for row in ws.iter_rows():
            for cell in row:
                try:
                    orig_val = orig_ws.cell(row=cell.row, column=cell.column).value
                except (KeyError, ValueError):
                    orig_val = None

                if str(cell.value) != str(orig_val):
                    # Mark with green fill + red border
                    cell.fill = GREEN_FILL
                    cell.border = RED_BORDER
                    changes += 1

    orig_wb.close()
    return changes


# ============================================================
#  4. Export .bas module into .xlsm workbook (Windows only, via COM)
# ============================================================
def export_vba_module(xlsm_path: str, bas_path: str) -> bool:
    """
    Embed a .bas VBA module into an existing .xlsm workbook using COM.

    Parameters
    ----------
    xlsm_path : str
        Path to the .xlsm workbook.
    bas_path : str
        Path to the .bas file to import.

    Returns
    -------
    bool — True if successful, False otherwise.
    """
    if not os.path.exists(bas_path):
        print(f"Error: .bas file not found: {bas_path}")
        return False

    # Ensure output is .xlsm
    if not xlsm_path.lower().endswith(".xlsm"):
        print("Warning: output should be .xlsm for macro support. Renaming.")
        xlsm_path = xlsm_path.rsplit(".", 1)[0] + ".xlsm"

    # Create a fresh .xlsm if it doesn't exist
    if not os.path.exists(xlsm_path):
        wb = Workbook()
        wb.save(xlsm_path)

    # Attempt COM import
    try:
        import win32com.client  # type: ignore

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb_com = excel.Workbooks.Open(os.path.abspath(xlsm_path))
        module_name = os.path.splitext(os.path.basename(bas_path))[0]

        # Remove existing module if present
        try:
            wb_com.VBProject.VBComponents(module_name).Activate
            wb_com.VBProject.VBComponents.Remove(
                wb_com.VBProject.VBComponents(module_name)
            )
        except Exception:
            pass

        # Import .bas file
        wb_com.VBProject.VBComponents.Import(os.path.abspath(bas_path))
        wb_com.Save()
        wb_com.Close()
        excel.Quit()

        print(f"Successfully imported {os.path.basename(bas_path)} into {xlsm_path}")
        return True

    except ImportError:
        print("pywin32 not installed. Install with: pip install pywin32")
        print(f"Manually import {bas_path} into {xlsm_path} via Excel VBA Editor.")
        return False
    except Exception as e:
        print(f"COM import failed: {e}")
        print("Ensure Excel 'Trust access to the VBA project object model' is enabled.")
        print(f"Manually import {bas_path} into {xlsm_path} via Excel VBA Editor.")
        return False


# ============================================================
#  Main entry point
# ============================================================
if __name__ == "__main__":
    print("CDM Excel Toolkit Plus — use from your own scripts:")
    print("  from cdm_toolkit_plus import *")
    print("  generate_file_listing('C:/Data', 'listing.xlsx')")
    print("  generate_folder_tree('C:/Data', 'tree.xlsx')")
    print("  track_changes_with_comments(wb, 'original.xlsx')")
    print("  export_vba_module('workbook.xlsm', 'module.bas')")
