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
# 19. Strikethrough Row Management
# ============================================================
def delete_strikethrough_rows(ws: openpyxl.worksheet.worksheet.Worksheet) -> int:
    """
    Delete rows where any cell has strikethrough formatting.
    Works bottom-up to preserve row indices during deletion.

    Returns
    -------
    int — number of rows deleted
    """
    rows_to_delete: set[int] = set()
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            if cell.value is not None and cell.font and cell.font.strikethrough:
                rows_to_delete.add(cell.row)
                break

    # Delete bottom-up
    deleted = 0
    for r in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(r)
        deleted += 1

    return deleted


def clear_strikethrough(ws: openpyxl.worksheet.worksheet.Worksheet) -> int:
    """
    Remove strikethrough formatting from all cells.

    Returns
    -------
    int — number of cells cleared
    """
    cleared = 0
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            if cell.font and cell.font.strikethrough:
                cell.font = Font(
                    name=cell.font.name,
                    size=cell.font.size,
                    bold=cell.font.bold,
                    italic=cell.font.italic,
                    underline=cell.font.underline,
                    color=cell.font.color,
                    strikethrough=False,
                )
                cleared += 1
    return cleared


# ============================================================
# 20. Sheet Creation / Deletion
# ============================================================
def create_sheets_from_list(wb: Workbook,
                            sheet_names: list[str],
                            clear_first: bool = False) -> int:
    """
    Batch create worksheets from a name list.
    Handles: trim to 31 chars, replace illegal filename chars,
    handle duplicates with _1/_2 suffix.

    Parameters
    ----------
    wb : Workbook
    sheet_names : list[str]
    clear_first : bool
        If True, delete existing sheets with same names before creating.

    Returns
    -------
    int — number of sheets created
    """
    import re
    created = 0
    seen: set[str] = set()

    for name in sheet_names:
        name = str(name).strip()
        if not name:
            continue

        # Trim to 31 chars (Excel sheet name limit)
        name = name[:31]
        # Replace illegal characters
        name = re.sub(r'[\\/*?:\[\]]', '_', name)

        # Handle duplicates
        base_name = name
        suffix = 1
        while name.lower() in seen:
            suffix_str = f"_{suffix}"
            max_base = 31 - len(suffix_str)
            name = base_name[:max_base] + suffix_str
            suffix += 1

        seen.add(name.lower())

        # Delete existing if requested
        if clear_first and name in wb.sheetnames:
            del wb[name]

        # Skip if already exists (and clear_first is False)
        if name in wb.sheetnames:
            continue

        wb.create_sheet(name)
        created += 1

    return created


def delete_sheets(wb: Workbook, sheet_names: list[str]) -> int:
    """
    Delete specified sheets by name. Cannot delete the last visible sheet.

    Returns
    -------
    int — number of sheets deleted
    """
    deleted = 0
    visible = [s for s in wb.sheetnames if wb[s].sheet_state == 'visible']

    for name in sheet_names:
        name = str(name).strip()
        if name in wb.sheetnames:
            # Never delete the last visible sheet
            if len(visible) <= 1 and wb[name].sheet_state == 'visible':
                continue
            if wb[name].sheet_state == 'visible':
                visible.remove(name)
            del wb[name]
            deleted += 1

    return deleted


# ============================================================
# 21. Medical Calculators — CrCl & eGFR
# ============================================================
def calc_cockcroft_gault(age: float, weight_kg: float, scr_mg_dl: float,
                         is_female: bool = False) -> float:
    """
    Cockcroft-Gault Creatinine Clearance.

    Formula: CrCl = ((140 - age) * weight_kg) / (72 * Scr_mg_dL) * 0.85 (if female)

    Returns -1 if Scr <= 0.
    """
    if scr_mg_dl <= 0:
        return -1.0
    crcl = ((140 - age) * weight_kg) / (72.0 * scr_mg_dl)
    if is_female:
        crcl *= 0.85
    return round(crcl, 2)


def calc_ckd_epi(age: float, scr_mg_dl: float,
                 is_female: bool = False) -> float:
    """
    CKD-EPI 2021 eGFR formula.

    eGFR = 142 * (Scr/A)^B * 0.9938^age * (1.012 if female)

    Where:
      A = 0.7 (female) / 0.9 (male)
      B = -0.241 (if Scr <= A*female_factor) else different values per gender

    Returns -1 if Scr <= 0.
    """
    if scr_mg_dl <= 0:
        return -1.0

    if is_female:
        a = 0.7
        # B depends on Scr relative to A
        if scr_mg_dl <= 0.7:
            b = -0.241
        else:
            b = -1.2
    else:
        a = 0.9
        if scr_mg_dl <= 0.9:
            b = -0.302
        else:
            b = -1.2

    egfr = 142.0 * ((scr_mg_dl / a) ** b) * (0.9938 ** age)
    if is_female:
        egfr *= 1.012

    return round(egfr, 2)


def calc_mdrd_egfr(age: float, scr_mg_dl: float,
                   is_female: bool = False, is_black: bool = False) -> float:
    """
    MDRD Study Equation for eGFR (original 4-variable formula).

    eGFR = 175 * Scr^(-1.154) * Age^(-0.203) * 0.742(if female) * 1.212(if black)

    Returns -1 if Scr <= 0.
    """
    if scr_mg_dl <= 0:
        return -1.0

    egfr = 175.0 * (scr_mg_dl ** -1.154) * (age ** -0.203)
    if is_female:
        egfr *= 0.742
    if is_black:
        egfr *= 1.212

    return round(egfr, 2)


def calc_bmi(weight_kg: float, height_cm: float) -> float:
    """
    Body Mass Index (BMI).

    BMI = weight(kg) / height(m)²

    Returns -1 if height <= 0.
    """
    if height_cm <= 0:
        return -1.0
    height_m = height_cm / 100.0
    return round(weight_kg / (height_m ** 2), 1)


def calc_bsa_mosteller(weight_kg: float, height_cm: float) -> float:
    """
    Body Surface Area — Mosteller formula.

    BSA (m²) = sqrt((height_cm * weight_kg) / 3600)

    Returns -1 if inputs <= 0.
    """
    import math
    if height_cm <= 0 or weight_kg <= 0:
        return -1.0
    return round(math.sqrt((height_cm * weight_kg) / 3600), 2)


def calc_ldl_friedewald(tc: float, hdl: float, tg: float) -> float:
    """
    LDL Cholesterol — Friedewald formula (mmol/L).

    LDL = TC - HDL - TG/2.2

    Note: only valid when TG < 4.5 mmol/L.
    Returns -1 if TG >= 4.5 or inputs invalid.
    """
    if tg >= 4.5 or tc <= 0 or hdl < 0 or tg <= 0:
        return -1.0
    return round(tc - hdl - (tg / 2.2), 1)


def calc_corrected_calcium(ca: float, albumin: float,
                           normal_albumin: float = 4.0) -> float:
    """
    Corrected Calcium for hypoalbuminemia.

    Corrected Ca = measured Ca + 0.8 * (normal_albumin - measured_albumin)

    Default normal albumin = 4.0 g/dL.
    """
    return round(ca + 0.8 * (normal_albumin - albumin), 1)


def calc_ibw_devine(height_cm: float, is_female: bool = False) -> float:
    """
    Ideal Body Weight — Devine formula.

    Male:   IBW(kg) = 50.0 + 2.3 * (height_inches - 60)
    Female: IBW(kg) = 45.5 + 2.3 * (height_inches - 60)

    Returns -1 if height <= 0.
    """
    if height_cm <= 0:
        return -1.0
    height_in = height_cm / 2.54
    base = 45.5 if is_female else 50.0
    return round(base + 2.3 * (height_in - 60), 1)


def calc_anion_gap(na: float, cl: float, hco3: float) -> float:
    """
    Anion Gap.

    AG = Na⁺ - (Cl⁻ + HCO₃⁻)

    Normal range: 8-12 mEq/L.
    """
    return round(na - (cl + hco3), 1)


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


# ============================================================
# 22. SAS Dataset Charting
# ============================================================
def sas_to_dataframe(sas_path: str) -> "pd.DataFrame":
    """Read a .sas7bdat file into a pandas DataFrame. Tries multiple encodings."""
    import pandas as pd

    # Check if pyreadstat is available (better SAS reader)
    try:
        import pyreadstat as _prs
        _has_pyreadstat = True
    except ImportError:
        _has_pyreadstat = False

    # Encodings to try (Chinese SAS data is often GBK, not UTF-8)
    ENCODINGS = ("utf-8", "gbk", "gb2312", "gb18030", "latin-1")

    if _has_pyreadstat:
        for enc in ENCODINGS:
            try:
                df, _ = _prs.read_sas7bdat(sas_path, encoding=enc)
                return df
            except Exception:
                continue  # pyreadstat raises ReadstatError, not UnicodeDecodeError

    # Fallback: pandas built-in SAS reader
    for enc in ENCODINGS:
        try:
            return pd.read_sas(sas_path, encoding=enc)
        except Exception:
            continue

    # Last resort
    return pd.read_sas(sas_path, encoding=None)


def generate_chart_from_sas(
    sas_path: str,
    output_path: str,
    chart_type: str = "bar",
    x_col: str = "",
    y_col: str = "",
    group_col: str = "",
    title: str = "SAS Data Chart",
    max_rows: int = 1000,
) -> str:
    """
    Read a SAS dataset and generate an Excel file with an embedded chart.

    Parameters
    ----------
    sas_path : str — path to .sas7bdat file
    output_path : str — where to save the .xlsx output
    chart_type : str — "bar", "line", "pie", "scatter"
    x_col : str — column for X axis / categories
    y_col : str — column for Y axis / values
    group_col : str — optional grouping column (creates multiple series)
    title : str — chart title
    max_rows : int — max rows to read (default 1000)

    Returns
    -------
    str — path to the output .xlsx file
    """
    import pandas as pd
    from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference
    from openpyxl.chart.series import DataPoint
    from openpyxl.utils import get_column_letter

    df = sas_to_dataframe(sas_path)
    if len(df) > max_rows:
        df = df.head(max_rows)

    wb = Workbook()
    ws_data = wb.active
    ws_data.title = "Data"

    # Select columns
    cols = list(df.columns)
    if x_col not in cols:
        x_col = cols[0] if len(cols) > 0 else ""
    numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
    if y_col not in numeric_cols:
        y_col = numeric_cols[0] if numeric_cols else (cols[1] if len(cols) > 1 else x_col)

    # Write headers + data
    headers = [x_col, y_col] if not group_col else [x_col, group_col, y_col]
    for ci, h in enumerate(headers, 1):
        ws_data.cell(row=1, column=ci, value=h)
        ws_data.cell(row=1, column=ci).font = HEADER_FONT
        ws_data.cell(row=1, column=ci).fill = HEADER_FILL

    for ri, (_, row) in enumerate(df.iterrows()):
        for ci, h in enumerate(headers, 1):
            val = row.get(h)
            if isinstance(val, float) and (pd.isna(val)):
                val = None
            ws_data.cell(row=ri + 2, column=ci, value=val)

    n_rows = len(df) + 1

    # Create chart
    if chart_type == "bar":
        chart = BarChart()
    elif chart_type == "line":
        chart = LineChart()
    elif chart_type == "pie":
        chart = PieChart()
    elif chart_type == "scatter":
        chart = ScatterChart()
    else:
        chart = BarChart()

    chart.title = title
    chart.style = 10
    chart.y_axis.title = y_col
    chart.x_axis.title = x_col

    if chart_type == "pie":
        data_ref = Reference(ws_data, min_col=2, min_row=1, max_row=n_rows)
        cats_ref = Reference(ws_data, min_col=1, min_row=2, max_row=n_rows)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
    else:
        if group_col and group_col in cols:
            # Multiple series by group
            groups = df[group_col].dropna().unique()
            col_offset = 3  # x, group, y → y starts at col 3
            for gi, g in enumerate(groups[:10]):  # max 10 series
                gdf = df[df[group_col] == g]
                for ri, (_, row) in enumerate(gdf.iterrows()):
                    ws_data.cell(row=ri + 2, column=col_offset + gi, value=row.get(y_col))
                ws_data.cell(row=1, column=col_offset + gi, value=str(g))
                ws_data.cell(row=1, column=col_offset + gi).font = HEADER_FONT
                ws_data.cell(row=1, column=col_offset + gi).fill = HEADER_FILL
                data_ref = Reference(ws_data, min_col=col_offset + gi, min_row=1, max_row=len(gdf) + 1)
                chart.add_data(data_ref, titles_from_data=True)
            cats_ref = Reference(ws_data, min_col=1, min_row=2, max_row=n_rows)
            chart.set_categories(cats_ref)
        else:
            y_col_idx = 2
            data_ref = Reference(ws_data, min_col=y_col_idx, min_row=1, max_row=n_rows)
            cats_ref = Reference(ws_data, min_col=1, min_row=2, max_row=n_rows)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)

    # Add chart to a new sheet
    ws_chart = wb.create_sheet("Chart")
    ws_chart.add_chart(chart, "B2")

    auto_width(ws_data)
    wb.save(output_path)
    return output_path


def sas_preview(sas_path: str, rows: int = 100) -> dict:
    """Preview a SAS dataset — return headers + first N rows as JSON."""
    df = sas_to_dataframe(sas_path)
    if len(df) > rows:
        df = df.head(rows)
    import pandas as pd
    headers = list(df.columns)
    data_rows = []
    for _, row in df.iterrows():
        data_rows.append([
            str(v)[:200] if not (isinstance(v, float) and pd.isna(v)) else ""
            for v in row
        ])
    return {
        "headers": headers,
        "rows": data_rows,
        "total_rows": len(df),
        "numeric_cols": df.select_dtypes(include=["number"]).columns.tolist(),
    }


# ============================================================
# 23. Raw Data QC Analysis
# ============================================================
def scan_sas_metadata(sas_path: str) -> dict:
    """
    Scan a SAS dataset and return metadata only (no QC findings).
    Pure data exploration — dataset info, variable listing, sample values.
    """
    import pandas as pd

    df = sas_to_dataframe(sas_path)
    n_rows = len(df)
    n_cols = len(df.columns)

    dataset_info = {
        "file_name": os.path.basename(sas_path),
        "row_count": n_rows,
        "col_count": n_cols,
        "memory_mb": round(df.memory_usage(deep=True).sum() / (1024 * 1024), 2),
    }

    variables = []
    for col in df.columns:
        dtype = str(df[col].dtype)
        missing = int(df[col].isna().sum())
        missing_pct = round(100 * missing / n_rows, 1) if n_rows > 0 else 0
        unique = int(df[col].nunique())
        sample_vals = df[col].dropna().head(3).tolist()
        sample_vals = [str(v)[:50] for v in sample_vals]
        var_type = "Char" if dtype == "object" else "Num"

        variables.append({
            "name": str(col),
            "type": var_type,
            "dtype": dtype,
            "missing": missing,
            "missing_pct": missing_pct,
            "unique": unique,
            "sample": sample_vals[:3],
        })

    char_count = sum(1 for v in variables if v["type"] == "Char")
    num_count = n_cols - char_count

    return {
        "dataset": dataset_info,
        "variables": variables,
        "summary": {
            "total_vars": n_cols,
            "char_vars": char_count,
            "num_vars": num_count,
            "complete_vars": sum(1 for v in variables if v["missing"] == 0),
        },
    }
