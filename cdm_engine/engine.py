"""
CDM Engine — Unified business logic for clinical data management Excel workflows.

Combines the functionality previously split across cdm_toolkit.py (14 functions)
and cdm_toolkit_plus.py (4 functions) into a single module that imports constants
from cdm_engine.conversions.

Features:
  1.  Text/term standardisation (find & replace batch)
  2.  Duplicate highlighting (consecutive equal cells → yellow)
  3.  Blank row insertion (every N rows or between group changes)
  4.  Blank row / hidden sheet removal
  5.  Table of Contents (TOC) generator with hyperlinks
  6.  SAS Derive-style TOC (description in B column)
  7.  Sheet splitter (split one sheet into many by column value)
  8.  Workbook merger (consolidate multiple .xlsx into one)
  9.  Track changes (highlight cell modifications)
  10. Changes report (structured dict output)
  11. Date calculator (add/subtract days from dates)
  12. Date difference calculator
  13. Unit conversion for lab data
  14. Batch sheet rename
  15. Special character scanner
  16. File directory listing → Excel with hyperlinks
  17. Folder tree → Excel with indentation
  18. Web API file-based wrapper functions
"""

from __future__ import annotations
import os
import re
import tempfile
from datetime import datetime, timedelta
from copy import copy

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.hyperlink import Hyperlink

from cdm_engine.conversions import (
    TERM_MAP, UNIT_CONVERSIONS,
    YELLOW_FILL, GREEN_FILL, RED_FILL, GREY_FILL, LIGHT_BLUE,
    HEADER_FILL, HEADER_FONT, TOC_FONT, BOLD_FONT,
    THIN_BORDER, RED_BORDER,
)


# ============================================================
#  1. Text / Term Standardisation
# ============================================================
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
#  4. Blank Row / Hidden Sheet Removal
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
#  5. Table of Contents (TOC) with Hyperlinks
# ============================================================
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
#  6. SAS Derive-Style TOC  (Description in column B)
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
#  7. Sheet Splitter — one column → many sheets
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
#  8. Workbook Merger
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
#  9. Track Changes (cell modification highlighting)
# ============================================================
def track_changes(ws: openpyxl.worksheet.worksheet.Worksheet,
                  original_wb_path: str):
    """
    Compare *ws* against the same sheet in *original_wb_path* and mark
    changed cells with a red border and green fill.

    Parameters
    ----------
    ws : Worksheet — the modified sheet
    original_wb_path : str — path to the original workbook for comparison

    Returns
    -------
    int — number of changed cells detected
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
                cell.border = RED_BORDER
                changes += 1
    orig_wb.close()
    return changes


# ============================================================
# 10. Changes Report (structured dict output)
# ============================================================
def changes_report(ws: openpyxl.worksheet.worksheet.Worksheet,
                   original_wb_path: str) -> list[dict]:
    """
    Compare *ws* against the same sheet in *original_wb_path* and return
    a structured list of changes without modifying the worksheet.

    Parameters
    ----------
    ws : Worksheet — the modified sheet
    original_wb_path : str — path to the original workbook for comparison

    Returns
    -------
    list[dict] — each dict has keys: address, type, old_value, new_value
    """
    orig_wb = load_workbook(original_wb_path, data_only=True)
    changes: list[dict] = []
    if ws.title not in orig_wb.sheetnames:
        orig_wb.close()
        return changes

    orig_ws = orig_wb[ws.title]
    for row in ws.iter_rows():
        for cell in row:
            orig_val = orig_ws.cell(row=cell.row, column=cell.column).value
            if cell.value != orig_val:
                changes.append({
                    'address': cell.coordinate,
                    'type': type(cell.value).__name__ if cell.value is not None else 'NoneType',
                    'old_value': str(orig_val) if orig_val is not None else '',
                    'new_value': str(cell.value) if cell.value is not None else '',
                })
    orig_wb.close()
    return changes


# ============================================================
# 11. Date Calculator
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
# 12. Lab Unit Conversion
# ============================================================
def convert_lab_units(ws: openpyxl.worksheet.worksheet.Worksheet,
                      test_col: str,       # column with lab test name
                      result_col: str,     # column with numeric result
                      unit_col: str,       # column with unit
                      standard_unit: str,  # target standard unit
                      test_filter: str | None = None):  # optional: only for this test
    """
    Convert lab results to a standard unit using built-in conversion factors.

    For each row where unit != standard_unit:
        result = result * factor

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
# 13. Batch Sheet Rename
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
# 14. Special Character Scanner
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
# 15. File Directory Listing → Excel
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
# 16. Folder Tree → Excel
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
# 17. Export / Batch Formatting Utilities
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
# 18. Web API File-Based Wrapper Functions
# ============================================================
def merge_workbooks_from_uploads(uploaded_files,
                                 create_toc: bool = True,
                                 add_prefix: bool = True) -> str:
    """
    Merge uploaded workbook files into one and return the path to a temporary
    merged file.  Suitable for web API endpoints that receive file uploads.

    Parameters
    ----------
    uploaded_files : list
        List of file paths (str), file-like objects, or FastAPI UploadFile objects.
    create_toc : bool
        Whether to add a TOC sheet (default True).
    add_prefix : bool
        Whether to prefix sheet names with source file name (default True).

    Returns
    -------
    str — path to the temporary merged .xlsx file
    """
    temp_dir = tempfile.mkdtemp()
    file_paths: list[str] = []

    for i, f in enumerate(uploaded_files):
        if isinstance(f, str):
            file_paths.append(f)
            continue

        # FastAPI UploadFile: read bytes from underlying .file
        content = None
        if hasattr(f, 'file') and hasattr(f.file, 'read'):
            content = f.file.read()
        elif hasattr(f, 'read') and callable(f.read):
            result = f.read()
            content = result if isinstance(result, bytes) else result.encode('utf-8')
        elif hasattr(f, 'read'):
            content = f.read

        if content is None:
            raise TypeError(f"Cannot handle uploaded file type: {type(f)}")

        # Save to temp, preserving original filename as prefix
        fname = getattr(f, 'filename', None) or f'_upload_{i}.xlsx'
        tmp_path = os.path.join(temp_dir, fname)
        if isinstance(content, str):
            content = content.encode('utf-8')
        with open(tmp_path, 'wb') as out:
            out.write(content)
        file_paths.append(tmp_path)

    output_path = os.path.join(temp_dir, 'merged_output.xlsx')
    merge_workbooks(file_paths, output_path, create_toc=create_toc)
    return output_path


def _upload_to_path(file_upload) -> str:
    """Convert a file upload (str path, file-like, or FastAPI UploadFile) to a temp file path."""
    if isinstance(file_upload, str):
        return file_upload

    # Extract bytes
    content = None
    if hasattr(file_upload, 'file') and hasattr(file_upload.file, 'read'):
        content = file_upload.file.read()
    elif hasattr(file_upload, 'read') and callable(file_upload.read):
        result = file_upload.read()
        content = result if isinstance(result, bytes) else result.encode('utf-8')
    elif hasattr(file_upload, 'read'):
        content = file_upload.read

    if content is None:
        raise TypeError(f"Cannot handle uploaded file type: {type(file_upload)}")
    if isinstance(content, str):
        content = content.encode('utf-8')

    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp.write(content)
    tmp.close()
    return tmp.name


def split_sheet_file(file_upload, split_col: str = "A",
                     header_rows: int = 1) -> str:
    """
    Load an uploaded workbook, split its first sheet by *split_col*, and save
    the result to a temporary file.
    """
    in_path = _upload_to_path(file_upload)
    wb = load_workbook(in_path)
    source_sheet = wb.sheetnames[0]
    split_sheet_by_column(wb, source_sheet, split_col=split_col,
                          header_rows=header_rows)
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp.close()
    wb.save(tmp.name)
    return tmp.name


def add_days_to_column_file(file_upload, col_letter: str, days: int,
                            date_format: str = "YYYY-MM-DD") -> str:
    """
    Load an uploaded workbook, add *days* to date values in *col_letter*
    on the active sheet, and save to a temporary file.
    """
    in_path = _upload_to_path(file_upload)
    wb = load_workbook(in_path)
    ws = wb.active
    add_days_to_column(ws, col_letter, days, row_start=2,
                       date_format=date_format)
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp.close()
    wb.save(tmp.name)
    return tmp.name


def convert_lab_units_file(file_upload, test_col: str, result_col: str,
                           unit_col: str, standard_unit: str,
                           test_filter: str | None = None) -> str:
    """
    Load an uploaded workbook, convert lab units on the active sheet, and save
    the result to a temporary file.

    Parameters
    ----------
    file_upload : str | file-like
        Path to an .xlsx file or a file-like object.
    test_col : str
        Column letter with lab test name.
    result_col : str
        Column letter with numeric result.
    unit_col : str
        Column letter with original unit.
    standard_unit : str
        Target standard unit, e.g. "mg".
    test_filter : str | None
        Optional test name filter.

    Returns
    -------
    str — path to the temporary output .xlsx file
    """
    in_path = _upload_to_path(file_upload)
    wb = load_workbook(in_path)
    ws = wb.active
    convert_lab_units(ws, test_col, result_col, unit_col, standard_unit,
                      test_filter=test_filter)
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp.close()
    wb.save(tmp.name)
    return tmp.name


def list_directory(path: str, include_subfolders: bool = True,
                   file_pattern: str = "*.*") -> dict:
    """
    Return a structured directory listing as a dict (for web API JSON responses).

    Parameters
    ----------
    path : str
        Root folder to scan.
    include_subfolders : bool
        Recursively scan subfolders (default True).
    file_pattern : str
        File filter, e.g. "*.xlsx" or "*.pdf" (default "*.*").

    Returns
    -------
    dict with keys: root (str), files (list[dict]), folders (list[dict])
    """
    if file_pattern in ("*.*", "*"):
        ext_filter = None
    elif file_pattern.startswith("*."):
        ext_filter = file_pattern[1:].lower()
    else:
        ext_filter = file_pattern.lower()

    items: list[dict] = []

    def _scan(scan_path: str) -> None:
        try:
            for entry in os.scandir(scan_path):
                if entry.is_file():
                    name = entry.name
                    if ext_filter is None or name.lower().endswith(ext_filter):
                        items.append({
                            'name': name,
                            'path': entry.path,
                            'is_dir': False,
                            'size': entry.stat().st_size,
                        })
                elif entry.is_dir():
                    items.append({
                        'name': entry.name,
                        'path': entry.path,
                        'is_dir': True,
                        'size': 0,
                    })
                    if include_subfolders:
                        _scan(entry.path)
        except PermissionError:
            pass

    _scan(path)
    return {'path': os.path.abspath(path), 'items': items, 'total': len(items)}


# ============================================================
#  Internal Helpers
# ============================================================
def _last_data_row(ws, col: int = 1) -> int:
    """Find the last non-empty row in *col*."""
    for r in range(ws.max_row, 0, -1):
        if ws.cell(row=r, column=col).value is not None:
            return r
    return 1


def _worksheet_exists(wb: Workbook, name: str) -> bool:
    return name in wb.sheetnames


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
