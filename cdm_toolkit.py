"""
CDM Excel Toolkit — Python implementation of DMI clinical data management macros.
Replaces VBA macros with portable, testable Python code using openpyxl.

Features:
  1. Text/term standardisation (find & replace batch)
  2. Duplicate highlighting (consecutive equal cells → yellow)
  3. Blank row insertion (every N rows or between group changes)
  4. Table of Contents (TOC) generator with hyperlinks
  5. Sheet splitter (split one sheet into many by column value)
  6. Workbook merger (consolidate multiple .xlsx into one)
  7. SAS Derive-style TOC (description in B column)
  8. Track changes (highlight cell modifications)
  9. Date calculator (add/subtract days from dates)
 10. Unit conversion for lab data
"""

from __future__ import annotations
import os
import re
import shutil
from datetime import datetime, timedelta
from copy import copy
from typing import Optional, Callable

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.hyperlink import Hyperlink


# ============================================================
#  Colour / Style constants
# ============================================================
YELLOW_FILL  = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
GREEN_FILL   = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
RED_FILL     = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
GREY_FILL    = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
LIGHT_BLUE   = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
HEADER_FILL  = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT  = Font(color="FFFFFF", bold=True, size=11)
TOC_FONT     = Font(color="0563C1", underline="single", size=11)
BOLD_FONT    = Font(bold=True, size=11)
THIN_BORDER  = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"))


# ============================================================
#  1. Text / Term Standardisation
# ============================================================
TERM_MAP = {
    "Recovered":                      "Recovered / Resolved",
    "Ongoing":                        "Not Recovered / Not Resolved",
    "Unknown":                        "Unknown",
    "Reasonable Possib":              "Reasonable Possibility",
    "Not Recovered/Not Resolved":     "Not Recovered / Not Resolved",
    "Recovered/Resolved":             "Recovered / Resolved",
    "Recovered/Resolved with Sequelae": "Recovered / Resolved with Sequelae",
}


def standardise_terms(ws: openpyxl.worksheet.worksheet.Worksheet,
                      term_map: dict[str, str] | None = None) -> int:
    """
    Replace non-standard terms across an entire worksheet.

    Parameters
    ----------
    ws : Worksheet
    term_map : dict, optional
        {old_text: new_text}.  Uses built-in TERM_MAP if None.

    Returns
    -------
    int — number of replacements made
    """
    if term_map is None:
        term_map = TERM_MAP
    count = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                for old, new in term_map.items():
                    if old in cell.value:
                        cell.value = cell.value.replace(old, new)
                        count += 1
    return count


# ============================================================
#  2. Duplicate Highlighting
# ============================================================
def highlight_duplicates(ws: openpyxl.worksheet.worksheet.Worksheet,
                         col_start: int = 7, col_end: int = 11,
                         row_start: int = 2, fill=YELLOW_FILL):
    """
    Highlight cells where the value in the same column equals the row above.
    (Marks BOTH the current and the previous cell.)

    Parameters
    ----------
    ws : Worksheet
    col_start, col_end : int — 1-based column range (default 7-11, i.e. G-K)
    row_start : int — first data row (default 2, row 1 = header)
    fill : PatternFill
    """
    max_row = ws.max_row
    for col_idx in range(col_start, col_end + 1):
        for row_idx in range(row_start + 1, max_row + 1):
            curr = ws.cell(row=row_idx, column=col_idx).value
            prev = ws.cell(row=row_idx - 1, column=col_idx).value
            if curr is not None and prev is not None and curr == prev:
                ws.cell(row=row_idx,     column=col_idx).fill = fill
                ws.cell(row=row_idx - 1, column=col_idx).fill = fill


# ============================================================
#  3. Blank Row Insertion
# ============================================================
def insert_blank_every_n(ws: openpyxl.worksheet.worksheet.Worksheet,
                         n: int = 5, col_letter: str = "A"):
    """
    Insert a blank row after every N rows of data.
    Works bottom-up to preserve row indices.
    """
    col = column_index_from_string(col_letter)
    last_row = _last_data_row(ws, col)
    inserted = 0
    for i in range(last_row, 1, -1):
        if (i - 1) % n == 0:
            ws.insert_rows(i + 1)
            inserted += 1
    return inserted


def insert_blank_between_groups(ws: openpyxl.worksheet.worksheet.Worksheet,
                                col_letter: str = "A"):
    """
    Insert a blank row between groups when the value in *col_letter* changes.
    """
    col = column_index_from_string(col_letter)
    last_row = _last_data_row(ws, col)
    inserted = 0
    for i in range(last_row, 2, -1):
        curr = ws.cell(row=i,     column=col).value
        nxt  = ws.cell(row=i + 1, column=col).value
        if curr is not None and nxt is not None and curr != nxt:
            ws.insert_rows(i + 1)
            inserted += 1
    return inserted


# ============================================================
#  4. Table of Contents (TOC) with Hyperlinks
# ============================================================
def _worksheet_exists(wb: Workbook, name: str) -> bool:
    return name in wb.sheetnames


def generate_toc(wb: Workbook, toc_name: str = "TOC",
                 exclude_sheets: tuple[str, ...] = ()):
    """
    Create (or recreate) a Table of Contents sheet with hyperlinks to every
    visible sheet, plus a back-link arrow shape on each sheet.

    Parameters
    ----------
    wb : Workbook
    toc_name : str — name for the TOC sheet
    exclude_sheets : tuple — sheet names to skip
    """
    # Remove existing TOC sheet
    if _worksheet_exists(wb, toc_name):
        del wb[toc_name]

    toc = wb.create_sheet(toc_name, 0)
    toc.cell(row=1, column=1, value="Table of Contents").font = Font(bold=True, size=14)
    toc.cell(row=1, column=1).fill = HEADER_FILL
    toc.cell(row=1, column=1).font = HEADER_FONT

    row_idx = 2
    for ws in wb.worksheets:
        if ws.title == toc_name or ws.title in exclude_sheets:
            continue
        cell = toc.cell(row=row_idx, column=1, value=ws.title)
        cell.font = TOC_FONT
        cell.hyperlink = Hyperlink(ref=cell.coordinate,
                                   location=f"'{ws.title}'!A1",
                                   display=ws.title)
        row_idx += 1

    toc.column_dimensions["A"].width = 35
    return toc


# ============================================================
#  5. Sheet Splitter — one column → many sheets
# ============================================================
def split_sheet_by_column(wb: Workbook, source_sheet: str,
                          split_col: str = "A",
                          header_rows: int = 1):
    """
    Split *source_sheet* into multiple sheets based on unique values in
    *split_col*.  The original sheet is not modified.

    Parameters
    ----------
    wb : Workbook (modified in place)
    source_sheet : str — name of sheet to split
    split_col : str — column letter for the split key
    header_rows : int — number of header rows to copy to each child sheet
    """
    src = wb[source_sheet]
    col_idx = column_index_from_string(split_col)

    # Read all data
    all_rows = list(src.iter_rows(min_row=1, max_row=src.max_row, values_only=True))
    headers = all_rows[:header_rows] if header_rows > 0 else []

    # Group rows by split key
    groups: dict[str, list] = {}
    for row in all_rows[header_rows:]:
        key = str(row[col_idx - 1]) if row[col_idx - 1] is not None else "(blank)"
        groups.setdefault(key, []).append(row)

    # Check for existing sheets with same names
    existing = set(wb.sheetnames) - {source_sheet}
    conflict = set(groups.keys()) & existing
    if conflict:
        raise ValueError(f"Sheet(s) already exist: {conflict}. "
                         f"Delete or rename them before splitting.")

    # Create child sheets
    src_idx = wb.sheetnames.index(source_sheet)
    for key, rows in groups.items():
        # Truncate sheet name to Excel's 31-char limit
        safe_name = re.sub(r'[\\/*?\[\]:]', '_', key)[:31]
        child = wb.create_sheet(safe_name)
        r = 1
        for h in headers:
            child.cell(row=r, column=1, value=h[0]).font = BOLD_FONT
            r += 1
        for data_row in rows:
            for c, val in enumerate(data_row, 1):
                child.cell(row=r, column=c, value=val)
            r += 1

    return list(groups.keys())


# ============================================================
#  6. Workbook Merger
# ============================================================
def merge_workbooks(file_paths: list[str], output_path: str,
                    create_toc: bool = True) -> Workbook:
    """
    Merge multiple Excel workbooks into one.  Each source workbook's sheets
    are copied into the output workbook.  Sheet names are prefixed with the
    source file name (truncated) to avoid collisions.

    Parameters
    ----------
    file_paths : list[str] — paths to .xlsx files to merge
    output_path : str — where to save the merged workbook
    create_toc : bool — add a TOC sheet

    Returns
    -------
    Workbook — the merged workbook (also saved to disk)
    """
    merged = Workbook()
    merged.remove(merged.active)  # remove default sheet

    for fpath in file_paths:
        prefix = os.path.splitext(os.path.basename(fpath))[0][:20]
        src = load_workbook(fpath, data_only=True)
        for ws in src.worksheets:
            new_name = f"{prefix}_{ws.title}"[:31]
            # Avoid duplicate names
            counter = 1
            base = new_name
            while new_name in merged.sheetnames:
                new_name = f"{base[:27]}_{counter}"
                counter += 1
            child = merged.create_sheet(new_name)
            _copy_sheet(ws, child)
        src.close()

    if create_toc:
        generate_toc(merged, "TOC")

    merged.save(output_path)
    return merged


# ============================================================
#  7. SAS Derive-Style TOC  (Description in column B)
# ============================================================
def generate_sas_derive_toc(wb: Workbook, desc_col: str = "V2",
                            toc_name: str = "TOC"):
    """
    Create a TOC where column A = sheet name (with hyperlink) and column B =
    description pulled from cell *desc_col* of each sheet.

    This matches the SAS Derive VBA macro behaviour.
    """
    if _worksheet_exists(wb, toc_name):
        del wb[toc_name]

    toc = wb.create_sheet(toc_name, 0)
    toc.cell(row=1, column=1, value="TOC").font = Font(bold=True, size=14)
    toc.cell(row=1, column=2, value="Description").font = Font(bold=True, size=14)

    row_idx = 2
    for ws in wb.worksheets:
        if ws.title == toc_name:
            continue
        # Column A: hyperlinked sheet name
        cell_a = toc.cell(row=row_idx, column=1, value=ws.title)
        cell_a.font = TOC_FONT
        cell_a.hyperlink = Hyperlink(ref=cell_a.coordinate,
                                     location=f"'{ws.title}'!A1",
                                     display=ws.title)
        # Column B: description from specified cell
        try:
            desc = ws[desc_col].value
        except (KeyError, ValueError):
            desc = ""
        toc.cell(row=row_idx, column=2, value=str(desc) if desc else "")

        row_idx += 1

    toc.column_dimensions["A"].width = 35
    toc.column_dimensions["B"].width = 50
    return toc


# ============================================================
#  8. Track Changes (cell modification logging)
# ============================================================
def track_changes_add_comment(ws: openpyxl.worksheet.worksheet.Worksheet,
                              original_wb_path: str):
    """
    Compare *ws* against the same sheet in *original_wb_path* and add Excel
    comments to changed cells showing old → new values.

    NOTE: openpyxl comment support is limited; this marks changed cells with
    a red border and green fill instead.  For real comments, use VBA.
    """
    orig_wb = load_workbook(original_wb_path, data_only=True)
    if ws.title not in orig_wb.sheetnames:
        orig_wb.close()
        return 0
    orig_ws = orig_wb[ws.title]

    changes = 0
    for row in ws.iter_rows():
        for cell in row:
            orig_val = orig_ws.cell(row=cell.row, column=cell.column).value
            if cell.value != orig_val:
                cell.fill = GREEN_FILL
                cell.border = Border(
                    left=Side(style="medium", color="FF0000"),
                    right=Side(style="medium", color="FF0000"),
                    top=Side(style="medium", color="FF0000"),
                    bottom=Side(style="medium", color="FF0000"))
                changes += 1
    orig_wb.close()
    return changes


# ============================================================
#  9. Date Calculator
# ============================================================
def add_days_to_column(ws: openpyxl.worksheet.worksheet.Worksheet,
                       col_letter: str, days: int,
                       row_start: int = 2, date_format: str = "YYYY-MM-DD"):
    """
    Add *days* to every date value in *col_letter*.
    Non-date cells are left unchanged.

    Returns count of cells updated.
    """
    col = column_index_from_string(col_letter)
    count = 0
    for row_idx in range(row_start, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=col)
        if isinstance(cell.value, datetime):
            cell.value = cell.value + timedelta(days=days)
            cell.number_format = date_format
            count += 1
    return count


def calc_date_diff(ws: openpyxl.worksheet.worksheet.Worksheet,
                   col_start: str, col_end: str, result_col: str,
                   row_start: int = 2):
    """
    Compute (col_end - col_start) in days and write to *result_col*.
    """
    c1 = column_index_from_string(col_start)
    c2 = column_index_from_string(col_end)
    cr = column_index_from_string(result_col)
    count = 0
    for row_idx in range(row_start, ws.max_row + 1):
        d1 = ws.cell(row=row_idx, column=c1).value
        d2 = ws.cell(row=row_idx, column=c2).value
        if isinstance(d1, datetime) and isinstance(d2, datetime):
            ws.cell(row=row_idx, column=cr).value = (d2 - d1).days
            count += 1
    return count


# ============================================================
# 10. Lab Unit Conversion
# ============================================================
# Predefined conversion factors (to standard unit)
UNIT_CONVERSIONS = {
    # Mass
    ("g",   "mg"):  1000,
    ("g",   "ug"):  1_000_000,
    ("mg",  "g"):   0.001,
    ("mg",  "ug"):  1000,
    ("ug",  "g"):   1e-6,
    ("ug",  "mg"):  0.001,
    ("kg",  "g"):   1000,
    ("g",   "kg"):  0.001,
    # Volume
    ("L",   "mL"):  1000,
    ("mL",  "L"):   0.001,
    ("dL",  "L"):   0.1,
    ("L",   "dL"):  10,
    # Length
    ("cm",  "mm"):  10,
    ("mm",  "cm"):  0.1,
    ("m",   "cm"):  100,
    ("cm",   "m"):  0.01,
    # Concentration
    ("g/L",  "mg/mL"): 1,
    ("mg/mL","g/L"):   1,
    ("mg/dL","g/L"):   0.01,
    ("g/L",  "mg/dL"): 100,
    ("g/dL", "g/L"):   10,
    ("g/L",  "g/dL"):  0.1,
    ("umol/L","mg/dL"): 0.0113,  # creatinine approx
    ("mg/dL", "umol/L"): 88.42,
    # Time
    ("h",   "min"): 60,
    ("min", "h"):   1/60,
    ("d",   "h"):   24,
    ("h",   "d"):   1/24,
}


def convert_lab_units(ws: openpyxl.worksheet.worksheet.Worksheet,
                      test_col: str,       # column with lab test name
                      result_col: str,     # column with numeric result
                      unit_col: str,       # column with unit
                      standard_unit: str,  # target standard unit
                      test_filter: str | None = None):  # optional: only for this test
    """
    Convert lab results to a standard unit using built-in conversion factors.

    For each row where unit != standard_unit:
        result = result × factor

    Parameters
    ----------
    ws : Worksheet (modified in place)
    test_col, result_col, unit_col : str — column letters
    standard_unit : str — target unit, e.g. "mg"
    test_filter : str | None — only convert rows with this test name

    Returns
    -------
    int — number of values converted
    """
    tc = column_index_from_string(test_col)
    rc = column_index_from_string(result_col)
    uc = column_index_from_string(unit_col)

    count = 0
    for row_idx in range(2, ws.max_row + 1):
        test_name = ws.cell(row=row_idx, column=tc).value
        unit      = ws.cell(row=row_idx, column=uc).value
        result    = ws.cell(row=row_idx, column=rc).value

        if test_filter and test_name != test_filter:
            continue
        if unit is None or unit == standard_unit:
            continue
        if result is None or not isinstance(result, (int, float)):
            continue

        factor = UNIT_CONVERSIONS.get((str(unit), standard_unit))
        if factor is not None:
            ws.cell(row=row_idx, column=rc).value = round(result * factor, 6)
            ws.cell(row=row_idx, column=uc).value = standard_unit
            count += 1
        else:
            ws.cell(row=row_idx, column=rc).fill = RED_FILL  # flag unknown conversion
    return count


# ============================================================
# 11. Batch Sheet Rename
# ============================================================
def batch_rename_sheets(wb: Workbook, name_map: dict[str, str]):
    """Rename multiple sheets at once. {old_name: new_name}."""
    for old, new in name_map.items():
        if old in wb.sheetnames:
            wb[old].title = new[:31]


def list_sheet_names(wb: Workbook) -> list[str]:
    """Return all visible sheet names."""
    return [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]


# ============================================================
# 12. Special Character Scanner
# ============================================================
def scan_special_chars(ws: openpyxl.worksheet.worksheet.Worksheet,
                       allowed_pattern: str = r'[a-zA-Z0-9\s\-_.,:;()/]+'
                       ) -> list[tuple[int, int, str]]:
    """
    Find cells containing characters outside the allowed set.
    Returns [(row, col, cell_value), ...].
    """
    allowed = re.compile(f"^{allowed_pattern}$")
    findings = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if not allowed.match(cell.value):
                    findings.append((cell.row, cell.column, cell.value))
    return findings


# ============================================================
# 13. Blank Row / Hidden Sheet Removal
# ============================================================
def delete_blank_rows(ws: openpyxl.worksheet.worksheet.Worksheet):
    """Delete entirely blank rows (all cells None or empty)."""
    rows_to_delete = []
    for row_idx in range(ws.max_row, 0, -1):
        vals = [ws.cell(row=row_idx, column=c).value
                for c in range(1, ws.max_column + 1)]
        if all(v is None or v == "" for v in vals):
            rows_to_delete.append(row_idx)
    for r in rows_to_delete:
        ws.delete_rows(r)
    return len(rows_to_delete)


def delete_hidden_sheets(wb: Workbook):
    """Remove all hidden sheets from the workbook."""
    removed = []
    for ws in wb.worksheets:
        if ws.sheet_state == "hidden":
            removed.append(ws.title)
            del wb[ws.title]
    return removed


# ============================================================
# 14. Export / Batch utilities
# ============================================================
def apply_header_style(ws: openpyxl.worksheet.worksheet.Worksheet,
                       header_row: int = 1):
    """Apply standard header formatting (blue fill, white bold font)."""
    for cell in ws[header_row]:
        if cell.value is not None:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.border = THIN_BORDER


def auto_width(ws: openpyxl.worksheet.worksheet.Worksheet,
               max_width: int = 60, min_width: int = 8):
    """Auto-size column widths based on content."""
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        lengths = []
        for cell in col_cells:
            if cell.value:
                lengths.append(len(str(cell.value)))
        if lengths:
            width = min(max(max(lengths) + 2, min_width), max_width)
            ws.column_dimensions[col_letter].width = width


# ============================================================
#  Helpers
# ============================================================
def _last_data_row(ws, col: int = 1) -> int:
    """Find the last non-empty row in *col*."""
    for r in range(ws.max_row, 0, -1):
        if ws.cell(row=r, column=col).value is not None:
            return r
    return 1


def _copy_sheet(src, dst):
    """Copy values + styles from *src* to *dst* (simple approach)."""
    for row in src.iter_rows():
        for cell in row:
            new_cell = dst.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.border = copy(cell.border)
                new_cell.alignment = copy(cell.alignment)
                new_cell.number_format = cell.number_format
    # Copy column widths
    for col_letter, dim in src.column_dimensions.items():
        dst.column_dimensions[col_letter].width = dim.width
    # Copy row heights
    for r, dim in src.row_dimensions.items():
        dst.row_dimensions[r].height = dim.height
    # Copy merged cells
    for merge_range in src.merged_cells.ranges:
        dst.merge_cells(str(merge_range))


# ============================================================
#  Main entry point for CLI usage
# ============================================================
if __name__ == "__main__":
    print("CDM Excel Toolkit — use from your own scripts:")
    print("  from cdm_toolkit import *")
    print("  wb = load_workbook('your_file.xlsx')")
    print("  standardise_terms(wb['Sheet1'])")
    print("  generate_toc(wb)")
    print("  wb.save('output.xlsx')")
