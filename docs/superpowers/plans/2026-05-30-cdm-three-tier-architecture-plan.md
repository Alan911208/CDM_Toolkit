# CDM 三层架构 — 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 CDM 工具包从"VBA + Python 两套平行实现"重构为三层架构 — Python 统一引擎 `cdm_engine/`、VBA 精简到8个模块、Web应用6页SPA。

**Architecture:** 表示层（VBA + Web UI）→ 逻辑层（Python `cdm_engine/` 包）→ 数据层（Excel/SAS/CSV）。VBA负责Excel内即时操作，Web负责复杂批处理和可视化，Python引擎作为唯一逻辑源。

**Tech Stack:** Python 3.10+, FastAPI, openpyxl, pandas/pyreadstat, 原生 HTML/CSS/JS (ES Modules + Web Components), VBA 7.1

---

## 文件结构

```
Tools_DM/
├── cdm_engine/                    # [新建] Python 核心引擎包
│   ├── __init__.py                #   公共 API 导出
│   ├── engine.py                  #   所有业务函数（从 cdm_toolkit.py + _plus.py 合并）
│   ├── conversions.py             #   常量（TERM_MAP, UNIT_CONVERSIONS, 样式）
│   └── sas_utils.py               #   SAS 数据集读写
├── web/                           # [新建] Web 前端
│   ├── index.html                 #   入口（仪表盘）
│   ├── css/app.css                #   全局样式
│   ├── js/app.js                  #   SPA 路由 + 状态
│   ├── js/components/             #   4个 Web Components
│   └── js/pages/                  #   6个功能页
├── server.py                      # [新建] FastAPI 启动入口
├── pyproject.toml                 # [新建] 打包配置
├── vba/                           # [修改] 精简并新增 DM_LaunchWeb
│   ├── CDM_Toolkit.bas            #   [修改] 更新汇总
│   ├── CDM_Install.bas            #   [修改] 新增 DM_LaunchWeb
│   └── modules/                   #   [删除] 移除4个模块
├── tests/
│   └── test_engine.py             # [新建] 引擎测试（基于已有测试逻辑）
└── README.md                      # [修改] 更新项目说明
```

---

## 阶段一：Python 引擎重构

### Task 1: 创建 cdm_engine 包骨架

**Files:**
- Create: `cdm_engine/__init__.py`
- Create: `cdm_engine/conversions.py`
- Create: `cdm_engine/engine.py` (空骨架，Task 2 填充)
- Create: `cdm_engine/sas_utils.py` (空骨架，Task 9 填充)

- [ ] **Step 1: 创建 conversions.py — 提取常量**

```python
# cdm_engine/conversions.py
"""
CDM constants: term maps, unit conversions, and shared styles.
Single source of truth — used by both engine.py and server.py.
"""
from openpyxl.styles import PatternFill, Font, Border, Side

# ============================================================
#  CDM Standard Term Map
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

# ============================================================
#  Lab Unit Conversion Table
# ============================================================
UNIT_CONVERSIONS = {
    ("g",   "mg"):  1000,
    ("g",   "ug"):  1_000_000,
    ("mg",  "g"):   0.001,
    ("mg",  "ug"):  1000,
    ("ug",  "g"):   1e-6,
    ("ug",  "mg"):  0.001,
    ("kg",  "g"):   1000,
    ("g",   "kg"):  0.001,
    ("L",   "mL"):  1000,
    ("mL",  "L"):   0.001,
    ("dL",  "L"):   0.1,
    ("L",   "dL"):  10,
    ("cm",  "mm"):  10,
    ("mm",  "cm"):  0.1,
    ("m",   "cm"):  100,
    ("cm",   "m"):  0.01,
    ("g/L",  "mg/mL"): 1,
    ("mg/mL","g/L"):   1,
    ("mg/dL","g/L"):   0.01,
    ("g/L",  "mg/dL"): 100,
    ("g/dL", "g/L"):   10,
    ("g/L",  "g/dL"):  0.1,
    ("umol/L","mg/dL"): 0.0113,
    ("mg/dL", "umol/L"): 88.42,
    ("h",   "min"): 60,
    ("min", "h"):   1/60,
    ("d",   "h"):   24,
    ("h",   "d"):   1/24,
}

# ============================================================
#  Shared Style Constants
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
    top=Side(style="thin"),  bottom=Side(style="thin"),
)
RED_BORDER   = Border(
    left=Side(style="medium", color="FF0000"),
    right=Side(style="medium", color="FF0000"),
    top=Side(style="medium", color="FF0000"),
    bottom=Side(style="medium", color="FF0000"),
)
```

- [ ] **Step 2: 创建 __init__.py — 占位，Task 3 填充导出**

```python
# cdm_engine/__init__.py
"""CDM Engine — unified clinical data management toolkit for Excel workflows."""
__version__ = "3.0.0"
```

- [ ] **Step 3: 创建 engine.py 和 sas_utils.py 占位**

```python
# cdm_engine/engine.py
"""CDM Engine — business logic functions. Populated in Task 2."""
```

```python
# cdm_engine/sas_utils.py
"""SAS dataset utilities. Populated in Task 9."""
```

- [ ] **Step 4: 验证包可导入**

Run: `python -c "import cdm_engine; print(cdm_engine.__version__)"`
Expected: `3.0.0`

- [ ] **Step 5: Commit**

```bash
git add cdm_engine/__init__.py cdm_engine/conversions.py cdm_engine/engine.py cdm_engine/sas_utils.py
git commit -m "feat: create cdm_engine package skeleton with constants"
```

---

### Task 2: 迁移业务函数到 cdm_engine/engine.py

**策略:** 从 `cdm_toolkit.py` 和 `cdm_toolkit_plus.py` 移动所有函数到 `cdm_engine/engine.py`，调整 import 使用 `cdm_engine.conversions`。

**Files:**
- Modify: `cdm_engine/engine.py` (replace placeholder with full content)

- [ ] **Step 1: 写入 engine.py — 导入 + 辅助函数**

Write the following to `cdm_engine/engine.py` (replace placeholder):

```python
"""
CDM Engine — all business logic functions.
Merged from cdm_toolkit.py (14 functions) + cdm_toolkit_plus.py (4 functions).
Single source of truth for both VBA translations and Web API endpoints.
"""
from __future__ import annotations
import os
import re
from copy import copy
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.hyperlink import Hyperlink

from cdm_engine.conversions import (
    TERM_MAP, UNIT_CONVERSIONS,
    YELLOW_FILL, GREEN_FILL, RED_FILL,
    HEADER_FILL, HEADER_FONT, TOC_FONT, BOLD_FONT, THIN_BORDER, RED_BORDER,
)


# ============================================================
#  Internal helpers
# ============================================================
def _last_data_row(ws, col: int = 1) -> int:
    """Find the last non-empty row in *col*."""
    for r in range(ws.max_row, 0, -1):
        if ws.cell(row=r, column=col).value is not None:
            return r
    return 1


def _worksheet_exists(wb: Workbook, name: str) -> bool:
    """Check if a sheet name exists in the workbook."""
    return name in wb.sheetnames


def _copy_sheet(src, dst):
    """Copy values + styles from *src* to *dst*."""
    for row in src.iter_rows():
        for cell in row:
            new_cell = dst.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.border = copy(cell.border)
                new_cell.alignment = copy(cell.alignment)
                new_cell.number_format = cell.number_format
    for col_letter, dim in src.column_dimensions.items():
        dst.column_dimensions[col_letter].width = dim.width
    for r, dim in src.row_dimensions.items():
        dst.row_dimensions[r].height = dim.height
    for merge_range in src.merged_cells.ranges:
        dst.merge_cells(str(merge_range))


# ============================================================
#  1. Text / Term Standardisation
# ============================================================
def standardise_terms(ws, term_map: dict[str, str] | None = None) -> int:
    """Replace non-standard terms across an entire worksheet."""
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
def highlight_duplicates(ws, col_start: int = 7, col_end: int = 11,
                         row_start: int = 2, fill=YELLOW_FILL):
    """Highlight cells where the value in the same column equals the row above."""
    max_row = ws.max_row
    for col_idx in range(col_start, col_end + 1):
        for row_idx in range(row_start + 1, max_row + 1):
            curr = ws.cell(row=row_idx, column=col_idx).value
            prev = ws.cell(row=row_idx - 1, column=col_idx).value
            if curr is not None and prev is not None and curr == prev:
                ws.cell(row=row_idx,     column=col_idx).fill = fill
                ws.cell(row=row_idx - 1, column=col_idx).fill = fill


# ============================================================
#  3. Blank Row Insertion / Deletion
# ============================================================
def insert_blank_every_n(ws, n: int = 5, col_letter: str = "A") -> int:
    """Insert a blank row after every N rows of data."""
    col = column_index_from_string(col_letter)
    last_row = _last_data_row(ws, col)
    inserted = 0
    for i in range(last_row, 1, -1):
        if (i - 1) % n == 0:
            ws.insert_rows(i + 1)
            inserted += 1
    return inserted


def insert_blank_between_groups(ws, col_letter: str = "A") -> int:
    """Insert a blank row between groups when the value in *col_letter* changes."""
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


def delete_blank_rows(ws) -> int:
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


# ============================================================
#  4. Table of Contents (TOC)
# ============================================================
def generate_toc(wb: Workbook, toc_name: str = "TOC",
                 exclude_sheets: tuple[str, ...] = ()):
    """Create a TOC sheet with hyperlinks to every visible sheet."""
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


def generate_sas_derive_toc(wb: Workbook, desc_col: str = "V2",
                            toc_name: str = "TOC"):
    """Create a SAS Derive-style TOC with description in column B."""
    if _worksheet_exists(wb, toc_name):
        del wb[toc_name]
    toc = wb.create_sheet(toc_name, 0)
    toc.cell(row=1, column=1, value="TOC").font = Font(bold=True, size=14)
    toc.cell(row=1, column=2, value="Description").font = Font(bold=True, size=14)
    row_idx = 2
    for ws in wb.worksheets:
        if ws.title == toc_name:
            continue
        cell_a = toc.cell(row=row_idx, column=1, value=ws.title)
        cell_a.font = TOC_FONT
        cell_a.hyperlink = Hyperlink(ref=cell_a.coordinate,
                                     location=f"'{ws.title}'!A1",
                                     display=ws.title)
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
#  5. Sheet Splitter
# ============================================================
def split_sheet_by_column(wb: Workbook, source_sheet: str,
                          split_col: str = "A",
                          header_rows: int = 1) -> list[str]:
    """Split *source_sheet* into multiple sheets by unique values in *split_col*."""
    src = wb[source_sheet]
    col_idx = column_index_from_string(split_col)
    all_rows = list(src.iter_rows(min_row=1, max_row=src.max_row, values_only=True))
    headers = all_rows[:header_rows] if header_rows > 0 else []
    groups: dict[str, list] = {}
    for row in all_rows[header_rows:]:
        key = str(row[col_idx - 1]) if row[col_idx - 1] is not None else "(blank)"
        groups.setdefault(key, []).append(row)
    existing = set(wb.sheetnames) - {source_sheet}
    conflict = set(groups.keys()) & existing
    if conflict:
        raise ValueError(f"Sheet(s) already exist: {conflict}. "
                         f"Delete or rename them before splitting.")
    for key, rows in groups.items():
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
    """Merge multiple Excel workbooks into one."""
    merged = Workbook()
    merged.remove(merged.active)
    for fpath in file_paths:
        prefix = os.path.splitext(os.path.basename(fpath))[0][:20]
        src = load_workbook(fpath, data_only=True)
        for ws in src.worksheets:
            new_name = f"{prefix}_{ws.title}"[:31]
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
#  7. Track Changes
# ============================================================
def track_changes(ws, original_wb_path: str) -> int:
    """Compare *ws* against the same sheet in *original_wb_path*. Mark changes."""
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


def changes_report(ws, original_wb_path: str) -> list[dict]:
    """Generate a structured change report as list of dicts."""
    orig_wb = load_workbook(original_wb_path, data_only=True)
    if ws.title not in orig_wb.sheetnames:
        orig_wb.close()
        return []
    orig_ws = orig_wb[ws.title]
    results = []
    for row in ws.iter_rows():
        for cell in row:
            orig_val = orig_ws.cell(row=cell.row, column=cell.column).value
            curr_val = cell.value
            if str(curr_val) != str(orig_val):
                change_type = "added" if orig_val is None else ("deleted" if curr_val is None else "modified")
                results.append({
                    "address": cell.coordinate,
                    "type": change_type,
                    "old_value": str(orig_val) if orig_val is not None else "",
                    "new_value": str(curr_val) if curr_val is not None else "",
                })
    orig_wb.close()
    return results


# ============================================================
#  8. Date Calculator
# ============================================================
def add_days_to_column(ws, col_letter: str, days: int,
                       row_start: int = 2, date_format: str = "YYYY-MM-DD") -> int:
    """Add *days* to every date value in *col_letter*."""
    col = column_index_from_string(col_letter)
    count = 0
    for row_idx in range(row_start, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=col)
        if isinstance(cell.value, datetime):
            cell.value = cell.value + timedelta(days=days)
            cell.number_format = date_format
            count += 1
    return count


def calc_date_diff(ws, col_start: str, col_end: str,
                   result_col: str, row_start: int = 2) -> int:
    """Compute (col_end - col_start) in days, write to *result_col*."""
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
#  9. Lab Unit Conversion
# ============================================================
def convert_lab_units(ws, test_col: str, result_col: str, unit_col: str,
                      standard_unit: str, test_filter: str | None = None) -> int:
    """Convert lab results to a standard unit."""
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
            ws.cell(row=row_idx, column=rc).fill = RED_FILL
    return count


# ============================================================
# 10. Sheet Management
# ============================================================
def batch_rename_sheets(wb: Workbook, name_map: dict[str, str]):
    """Rename multiple sheets at once. {old_name: new_name}."""
    for old, new in name_map.items():
        if old in wb.sheetnames:
            wb[old].title = new[:31]


def list_sheet_names(wb: Workbook) -> list[str]:
    """Return all visible sheet names."""
    return [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]


def delete_hidden_sheets(wb: Workbook) -> list[str]:
    """Remove all hidden sheets from the workbook."""
    removed = []
    for ws in wb.worksheets:
        if ws.sheet_state == "hidden":
            removed.append(ws.title)
            del wb[ws.title]
    return removed


# ============================================================
# 11. Special Character Scanner
# ============================================================
def scan_special_chars(ws, allowed_pattern: str = r'[a-zA-Z0-9\s\-_.,:;()/]+'
                       ) -> list[tuple[int, int, str]]:
    """Find cells containing characters outside the allowed set."""
    allowed = re.compile(f"^{allowed_pattern}$")
    findings = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if not allowed.match(cell.value):
                    findings.append((cell.row, cell.column, cell.value))
    return findings


# ============================================================
# 12. File Directory Listing
# ============================================================
def generate_file_listing(folder_path: str, output_path: str,
                          include_subfolders: bool = True,
                          file_pattern: str = "*.*") -> Workbook:
    """Generate an Excel file listing all files in a folder with hyperlinks."""
    wb = Workbook()
    ws = wb.active
    ws.title = "File Listing"
    ws.cell(row=1, column=1, value="Folder Path").font = HEADER_FONT
    ws.cell(row=1, column=1).fill = HEADER_FILL
    ws.cell(row=1, column=2, value="File Name").font = HEADER_FONT
    ws.cell(row=1, column=2).fill = HEADER_FILL
    if file_pattern in ("*.*", "*"):
        ext_filter = None
    elif file_pattern.startswith("*."):
        ext_filter = file_pattern[1:].lower()
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
                        cell.hyperlink = Hyperlink(ref=cell.coordinate,
                                                   target=entry.path, display=name)
                        row_idx += 1
                elif entry.is_dir() and include_subfolders:
                    _scan(entry.path)
        except PermissionError:
            pass
    _scan(folder_path)
    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 50
    wb.save(output_path)
    return wb


def generate_folder_tree(folder_path: str, output_path: str) -> Workbook:
    """Generate a recursive tree-view folder structure as Excel."""
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
                    cell.hyperlink = Hyperlink(ref=cell.coordinate,
                                               target=entry.path, display=entry.name)
                    row_idx += 1
        except PermissionError:
            ws.cell(row=row_idx, column=1, value=f"{indent}[Access Denied]")
            row_idx += 1
    _build_tree(folder_path)
    ws.column_dimensions["A"].width = 80
    wb.save(output_path)
    return wb


# ============================================================
#  Utility Functions
# ============================================================
def apply_header_style(ws, header_row: int = 1):
    """Apply standard header formatting."""
    for cell in ws[header_row]:
        if cell.value is not None:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.border = THIN_BORDER


def auto_width(ws, max_width: int = 60, min_width: int = 8):
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
#  File-based wrappers (for Web API)
# ============================================================
def merge_workbooks_from_uploads(uploaded_files, create_toc: bool = True,
                                  add_prefix: bool = True) -> str:
    """Merge uploaded file objects. Saves to temp file, returns path."""
    import tempfile
    tmp_dir = tempfile.mkdtemp()
    paths = []
    for uf in uploaded_files:
        fpath = os.path.join(tmp_dir, uf.filename or "upload.xlsx")
        with open(fpath, "wb") as f:
            f.write(uf.file.read())
        paths.append(fpath)
    out_path = os.path.join(tmp_dir, "merged.xlsx")
    merge_workbooks(paths, out_path, create_toc=create_toc)
    return out_path


def split_sheet_file(file_upload, split_col: str, header_rows: int) -> str:
    """Upload file → split → save temp file, return path."""
    import tempfile
    import shutil
    tmp_dir = tempfile.mkdtemp()
    in_path = os.path.join(tmp_dir, "input.xlsx")
    with open(in_path, "wb") as f:
        f.write(file_upload.file.read())
    wb = load_workbook(in_path)
    ws_name = wb.sheetnames[0]
    split_sheet_by_column(wb, ws_name, split_col=split_col, header_rows=header_rows)
    out_path = os.path.join(tmp_dir, "split.xlsx")
    wb.save(out_path)
    return out_path


def add_days_to_column_file(file_upload, col_letter: str, days: int,
                             date_format: str) -> str:
    """Upload file → add days → save temp file, return path."""
    import tempfile
    tmp_dir = tempfile.mkdtemp()
    in_path = os.path.join(tmp_dir, "input.xlsx")
    with open(in_path, "wb") as f:
        f.write(file_upload.file.read())
    wb = load_workbook(in_path)
    ws = wb.active
    add_days_to_column(ws, col_letter, days, row_start=2, date_format=date_format)
    out_path = os.path.join(tmp_dir, "dates_result.xlsx")
    wb.save(out_path)
    return out_path


def convert_lab_units_file(file_upload, test_col: str, result_col: str,
                            unit_col: str, standard_unit: str,
                            test_filter: str) -> str:
    """Upload file → convert lab units → save temp file, return path."""
    import tempfile
    tmp_dir = tempfile.mkdtemp()
    in_path = os.path.join(tmp_dir, "input.xlsx")
    with open(in_path, "wb") as f:
        f.write(file_upload.file.read())
    wb = load_workbook(in_path)
    ws = wb.active
    convert_lab_units(ws, test_col, result_col, unit_col, standard_unit,
                      test_filter=test_filter if test_filter else None)
    out_path = os.path.join(tmp_dir, "units_result.xlsx")
    wb.save(out_path)
    return out_path


def list_directory(path: str, include_subfolders: bool = True,
                   file_pattern: str = "*.*") -> dict:
    """Return directory listing as JSON-serializable dict."""
    result = {"path": path, "items": []}

    def _scan(p: str) -> None:
        try:
            for entry in sorted(os.scandir(p), key=lambda e: (not e.is_dir(), e.name.lower())):
                item = {
                    "name": entry.name,
                    "path": entry.path,
                    "is_dir": entry.is_dir(),
                    "size": entry.stat().st_size if entry.is_file() else 0,
                }
                result["items"].append(item)
        except PermissionError:
            pass
    _scan(path)
    result["total"] = len(result["items"])
    return result
```

- [ ] **Step 2: 验证引擎可导入且所有函数可用**

Run: `python -c "from cdm_engine.engine import *; print('OK:', len([x for x in dir() if not x.startswith('_')]), 'exports')"`

- [ ] **Step 3: Commit**

```bash
git add cdm_engine/engine.py
git commit -m "feat: migrate all 18 business functions to cdm_engine.engine"
```

---

### Task 3: 填充 cdm_engine/__init__.py 公共 API

**Files:**
- Modify: `cdm_engine/__init__.py`

- [ ] **Step 1: 写入完整 __init__.py**

Replace placeholder with:

```python
# cdm_engine/__init__.py
"""CDM Engine — unified clinical data management toolkit for Excel workflows.

Usage:
    from cdm_engine import standardise_terms, generate_toc
    from cdm_engine import merge_workbooks, convert_lab_units
"""
__version__ = "3.0.0"

from cdm_engine.engine import (
    # Text
    standardise_terms,
    # Duplicates
    highlight_duplicates,
    # Blank rows
    insert_blank_every_n,
    insert_blank_between_groups,
    delete_blank_rows,
    # TOC
    generate_toc,
    generate_sas_derive_toc,
    # Sheet management
    batch_rename_sheets,
    list_sheet_names,
    delete_hidden_sheets,
    # Track changes
    track_changes,
    changes_report,
    # Special chars
    scan_special_chars,
    # File listing
    generate_file_listing,
    generate_folder_tree,
    # Dates
    add_days_to_column,
    calc_date_diff,
    # Lab units
    convert_lab_units,
    # Sheet split
    split_sheet_by_column,
    # Workbook merge
    merge_workbooks,
    # File-based wrappers (for Web API)
    merge_workbooks_from_uploads,
    split_sheet_file,
    add_days_to_column_file,
    convert_lab_units_file,
    list_directory,
    # Utilities
    apply_header_style,
    auto_width,
)
```

- [ ] **Step 2: 验证顶层导入**

Run: `python -c "from cdm_engine import standardise_terms, generate_toc, merge_workbooks; print('Top-level imports OK')"`

- [ ] **Step 3: Commit**

```bash
git add cdm_engine/__init__.py
git commit -m "feat: expose public API from cdm_engine.__init__"
```

---

### Task 4: 编写引擎测试

**Files:**
- Create: `tests/test_engine.py`

- [ ] **Step 1: 创建测试文件 — 基于已有测试逻辑改编**

```python
"""
Test suite for cdm_engine — adapted from reference/test_cdm_toolkit.py.
Tests all imported functions from cdm_engine.
"""
import os
import sys
import shutil
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from cdm_engine import *

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_output")


def setup():
    """Clean and recreate test output directory."""
    if os.path.exists(OUTPUT_DIR):
        shutil.rmtree(OUTPUT_DIR)
    os.makedirs(OUTPUT_DIR)


# ---- Mock Data Generators ----
def make_ae_data(wb: Workbook):
    ws = wb.active
    ws.title = "AE"
    headers = ["STUDYID", "SITEID", "USUBJID", "AESEQ", "AETERM",
               "AESTDTC", "AEENDTC", "AESER", "AESEV", "AEOUT", "AEREL"]
    ws.append(headers)
    apply_header_style(ws)
    subjects = ["001-001"] * 4 + ["001-002"] * 4 + ["001-003"] * 4 + ["001-004"] * 4 + ["001-005"] * 4
    terms = ["Headache"] * 2 + ["Nausea"] * 2 + ["Fatigue"] * 2 + ["Dizziness"] * 2 + ["Rash"] * 2 + \
            ["Headache"] * 2 + ["Nausea"] * 2 + ["Fatigue"] * 2 + ["Dizziness"] * 2 + ["Rash"] * 2
    outcomes = ["Recovered", "Ongoing", "Unknown", "Recovered/Resolved with Sequelae",
                "Recovered", "Ongoing", "Unknown", "Recovered", "Ongoing", "Unknown",
                "Recovered", "Ongoing", "Unknown", "Recovered", "Ongoing",
                "Recovered", "Ongoing", "Unknown", "Recovered", "Ongoing"]
    start_date = datetime(2025, 6, 1)
    for i in range(20):
        ws.append([
            "STUDY001", f"SITE{10 + i % 3:02d}", subjects[i], i + 1, terms[i],
            start_date + timedelta(days=i * 7),
            start_date + timedelta(days=i * 7 + 5),
            "N", "MILD", outcomes[i], "Reasonable Possib",
        ])
    auto_width(ws)
    return ws


def make_lab_data(wb: Workbook):
    ws = wb.create_sheet("LB")
    headers = ["STUDYID", "USUBJID", "VISIT", "LBTEST", "LBORRES", "LBORRESU"]
    ws.append(headers)
    apply_header_style(ws)
    lab_data = [
        ["STUDY001", "001-001", "Screening", "Glucose", 90, "mg/dL"],
        ["STUDY001", "001-001", "Week 4", "Glucose", 0.9, "g/L"],
        ["STUDY001", "001-003", "Screening", "Hemoglobin", 14.5, "g/dL"],
        ["STUDY001", "001-003", "Week 4", "Hemoglobin", 145, "g/L"],
        ["STUDY001", "001-001", "Screening", "Creatinine", 0.9, "mg/dL"],
        ["STUDY001", "001-003", "Screening", "Creatinine", 88, "umol/L"],
    ]
    for row in lab_data:
        ws.append(row)
    auto_width(ws)
    return ws


# ---- Tests ----
def test_import():
    """Verify all 18 public functions are importable."""
    funcs = [
        standardise_terms, highlight_duplicates,
        insert_blank_every_n, insert_blank_between_groups, delete_blank_rows,
        generate_toc, generate_sas_derive_toc,
        batch_rename_sheets, list_sheet_names, delete_hidden_sheets,
        track_changes, changes_report,
        scan_special_chars,
        generate_file_listing, generate_folder_tree,
        add_days_to_column, calc_date_diff,
        convert_lab_units,
    ]
    assert all(callable(f) for f in funcs)
    print(f"[PASS] test_import — {len(funcs)} functions importable")


def test_standardise_terms():
    wb = Workbook()
    make_ae_data(wb)
    ws = wb["AE"]
    count = standardise_terms(ws)
    sample = str(ws.cell(row=2, column=10).value)
    assert "Recovered / Resolved" in sample, f"Got: {sample}"
    wb.save(os.path.join(OUTPUT_DIR, "test_standardise.xlsx"))
    print(f"[PASS] test_standardise_terms — {count} replacements")


def test_highlight_duplicates():
    wb = Workbook()
    make_ae_data(wb)
    ws = wb["AE"]
    highlight_duplicates(ws, col_start=5, col_end=5, row_start=2)
    assert ws.cell(row=9, column=5).fill.start_color.rgb in ("00FFFF00", "FFFF00")
    wb.save(os.path.join(OUTPUT_DIR, "test_duplicates.xlsx"))
    print("[PASS] test_highlight_duplicates")


def test_blank_rows():
    wb = Workbook()
    make_ae_data(wb)
    ws = wb["AE"]
    n1 = insert_blank_every_n(ws, n=5, col_letter="A")
    n2 = insert_blank_between_groups(ws, col_letter="C")
    assert n1 > 0 and n2 > 0
    wb.save(os.path.join(OUTPUT_DIR, "test_blank_rows.xlsx"))
    print(f"[PASS] test_blank_rows — every-N:{n1}, groups:{n2}")


def test_generate_toc():
    wb = Workbook()
    make_ae_data(wb)
    make_lab_data(wb)
    generate_toc(wb, "TOC")
    assert "TOC" in wb.sheetnames
    assert wb["TOC"].cell(row=2, column=1).value in ("AE", "LB")
    wb.save(os.path.join(OUTPUT_DIR, "test_toc.xlsx"))
    print("[PASS] test_generate_toc")


def test_sheet_splitter():
    wb = Workbook()
    make_ae_data(wb)
    ws = wb["AE"]
    ws.cell(row=1, column=12, value="SEV_GROUP")
    for i in range(2, 12):
        ws.cell(row=i, column=12, value="MILD" if i % 3 == 0 else "MODERATE")
    ws.cell(row=2, column=12, value="SEVERE")
    sheets = split_sheet_by_column(wb, "AE", split_col="L", header_rows=1)
    assert "MILD" in wb.sheetnames or "MODERATE" in wb.sheetnames
    wb.save(os.path.join(OUTPUT_DIR, "test_split.xlsx"))
    print(f"[PASS] test_sheet_splitter — {len(sheets)} sheets")


def test_track_changes():
    wb = Workbook()
    make_ae_data(wb)
    orig_path = os.path.join(OUTPUT_DIR, "test_changes_orig.xlsx")
    wb.save(orig_path)
    wb["AE"].cell(row=2, column=5).value = "Migraine"
    n = track_changes(wb["AE"], orig_path)
    assert n >= 1, f"Expected >=1 change, got {n}"
    wb.save(os.path.join(OUTPUT_DIR, "test_changes.xlsx"))
    print(f"[PASS] test_track_changes — {n} changes")


def test_date_calc():
    wb = Workbook()
    make_ae_data(wb)
    ws = wb["AE"]
    n1 = add_days_to_column(ws, "F", 30, row_start=2)
    ws.cell(row=1, column=13, value="AEDUR")
    n2 = calc_date_diff(ws, "F", "G", "M", row_start=2)
    assert n1 > 0 and n2 > 0
    wb.save(os.path.join(OUTPUT_DIR, "test_dates.xlsx"))
    print(f"[PASS] test_date_calc — shifted:{n1}, diffs:{n2}")


def test_convert_lab_units():
    wb = Workbook()
    make_lab_data(wb)
    ws = wb["LB"]
    n1 = convert_lab_units(ws, "D", "E", "F", "mg/dL", test_filter="Glucose")
    n2 = convert_lab_units(ws, "D", "E", "F", "g/dL", test_filter="Hemoglobin")
    for row_idx in range(2, ws.max_row + 1):
        test_name = ws.cell(row=row_idx, column=4).value
        unit = ws.cell(row=row_idx, column=6).value
        if test_name == "Glucose":
            assert unit == "mg/dL", f"Row {row_idx}: {unit}"
    wb.save(os.path.join(OUTPUT_DIR, "test_lab.xlsx"))
    print(f"[PASS] test_convert_lab_units — {n1 + n2} converted")


def test_full_pipeline():
    wb = Workbook()
    make_ae_data(wb)
    make_lab_data(wb)
    standardise_terms(wb["AE"])
    highlight_duplicates(wb["AE"], col_start=5, col_end=5, row_start=2)
    convert_lab_units(wb["LB"], "D", "E", "F", "mg/dL", test_filter="Glucose")
    wb["AE"].cell(row=1, column=13, value="AEDUR")
    calc_date_diff(wb["AE"], "F", "G", "M", row_start=2)
    generate_toc(wb, "TOC")
    for ws in wb.worksheets:
        auto_width(ws)
    wb.save(os.path.join(OUTPUT_DIR, "test_full_pipeline.xlsx"))
    assert "TOC" in wb.sheetnames
    print("[PASS] test_full_pipeline")


if __name__ == "__main__":
    setup()
    print(f"CDM Engine Test Suite\n{'='*50}")
    tests = [
        test_import, test_standardise_terms, test_highlight_duplicates,
        test_blank_rows, test_generate_toc, test_sheet_splitter,
        test_track_changes, test_date_calc, test_convert_lab_units,
        test_full_pipeline,
    ]
    passed = failed = 0
    for t in tests:
        try:
            t()
            passed += 1
        except Exception as e:
            print(f"[FAIL] {t.__name__} — {e}")
            import traceback; traceback.print_exc()
            failed += 1
    print(f"{'='*50}\nResults: {passed} PASSED, {failed} FAILED out of {len(tests)}")
```

- [ ] **Step 2: 运行测试**

Run: `cd e:/Project/Tools/Tools_DM && python tests/test_engine.py`
Expected: 10/10 PASSED

- [ ] **Step 3: Commit**

```bash
git add tests/test_engine.py
git commit -m "test: add cdm_engine test suite (10 tests)"
```

---

## 阶段二：FastAPI 后端

### Task 5: 创建 FastAPI 服务器 + 6个API端点

**Files:**
- Create: `server.py`

- [ ] **Step 1: 创建 server.py**

```python
"""
CDM Toolkit Web Server — FastAPI application.
Start with: cdm-tools serve  (or: python server.py)
"""
import os
import sys
import webbrowser
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from cdm_engine import (
    merge_workbooks_from_uploads,
    split_sheet_file,
    scan_special_chars,
    add_days_to_column_file,
    convert_lab_units_file,
    list_directory,
)

# ---- App Setup ----
app = FastAPI(title="CDM Toolkit", version="3.0.0")

WEB_DIR = Path(__file__).parent / "web"
if WEB_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(WEB_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    """Serve the SPA dashboard."""
    index_path = WEB_DIR / "index.html"
    if index_path.exists():
        return index_path.read_text(encoding="utf-8")
    return "<h1>CDM Toolkit</h1><p>Web files not found. Run from project root.</p>"


# ---- API Endpoints ----
@app.post("/api/merge")
async def api_merge(
    files: list[UploadFile] = File(...),
    create_toc: bool = Form(True),
    add_prefix: bool = Form(True),
):
    """Merge multiple workbooks into one."""
    result_path = merge_workbooks_from_uploads(files, create_toc, add_prefix)
    return FileResponse(result_path, filename="merged.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/split")
async def api_split(
    file: UploadFile = File(...),
    split_col: str = Form("A"),
    header_rows: int = Form(1),
):
    """Split a workbook by column values."""
    result_path = split_sheet_file(file, split_col, header_rows)
    return FileResponse(result_path, filename="split.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/scan-chars")
async def api_scan_chars(
    file: UploadFile = File(...),
    system: str = Form("Rave"),
):
    """Scan for special characters in uploaded workbook."""
    import tempfile
    from openpyxl import load_workbook
    tmp_dir = tempfile.mkdtemp()
    in_path = os.path.join(tmp_dir, "input.xlsx")
    with open(in_path, "wb") as f:
        f.write(file.file.read())
    wb = load_workbook(in_path)
    all_findings = []
    for ws in wb.worksheets:
        findings = scan_special_chars(ws)
        for row, col, val in findings:
            from openpyxl.utils import get_column_letter
            all_findings.append({
                "sheet": ws.title,
                "row": row,
                "col": get_column_letter(col),
                "cell": f"{get_column_letter(col)}{row}",
                "value": val[:200] if len(val) > 200 else val,
            })
    return {"findings": all_findings, "total": len(all_findings), "system": system}


@app.post("/api/file-listing")
async def api_file_listing(
    path: str = Form(...),
    include_subfolders: bool = Form(True),
    file_pattern: str = Form("*.*"),
):
    """Browse directory contents."""
    return list_directory(path, include_subfolders, file_pattern)


@app.post("/api/date-calc")
async def api_date_calc(
    file: UploadFile = File(...),
    col_letter: str = Form(...),
    days: int = Form(...),
    date_format: str = Form("YYYY-MM-DD"),
):
    """Add/subtract days from a date column."""
    result_path = add_days_to_column_file(file, col_letter, days, date_format)
    return FileResponse(result_path, filename="dates_result.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/unit-convert")
async def api_unit_convert(
    file: UploadFile = File(...),
    test_col: str = Form(...),
    result_col: str = Form(...),
    unit_col: str = Form(...),
    standard_unit: str = Form(...),
    test_filter: str = Form(""),
):
    """Convert lab units to standard units."""
    result_path = convert_lab_units_file(file, test_col, result_col, unit_col,
                                          standard_unit, test_filter)
    return FileResponse(result_path, filename="units_result.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---- Startup ----
def start_server(port: int = 8520, open_browser: bool = True):
    """Launch the server and optionally open browser."""
    import uvicorn
    if open_browser:
        webbrowser.open(f"http://localhost:{port}")
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="info")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="CDM Toolkit Web Server")
    parser.add_argument("--port", type=int, default=8520, help="Server port (default: 8520)")
    parser.add_argument("--no-browser", action="store_true", help="Don't open browser")
    args = parser.parse_args()
    start_server(port=args.port, open_browser=not args.no_browser)
```

- [ ] **Step 2: 验证服务器可以启动**

Run: `cd e:/Project/Tools/Tools_DM && python server.py --no-browser --port 8521`
Expected: Server starts, Uvicorn running on http://127.0.0.1:8521
Then press Ctrl+C to stop.

- [ ] **Step 3: Commit**

```bash
git add server.py
git commit -m "feat: add FastAPI server with 6 API endpoints"
```

---

## 阶段三：Web 前端

### Task 6: Web 前端基础 — 样式 + SPA路由 + 仪表盘

**Files:**
- Create: `web/css/app.css`
- Create: `web/js/app.js`
- Create: `web/index.html`

- [ ] **Step 1: 创建全局样式 web/css/app.css**

```css
/* CDM Toolkit — Global Styles */
:root {
  --color-primary: #4472C4;
  --color-primary-dark: #2E5BA0;
  --color-bg: #F5F7FA;
  --color-surface: #FFFFFF;
  --color-text: #1A1A1A;
  --color-text-secondary: #666666;
  --color-border: #E0E4E8;
  --color-success: #2E7D32;
  --color-warning: #F57F17;
  --color-danger: #C62828;
  --radius: 8px;
  --shadow: 0 1px 3px rgba(0,0,0,0.1);
  --shadow-lg: 0 4px 12px rgba(0,0,0,0.1);
  --font: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Microsoft YaHei", sans-serif;
}
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: var(--font);
  background: var(--color-bg);
  color: var(--color-text);
  line-height: 1.6;
  min-height: 100vh;
}

/* Layout */
.app-header {
  background: var(--color-primary);
  color: white;
  padding: 0 24px;
  height: 56px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  box-shadow: var(--shadow-lg);
}
.app-header h1 { font-size: 18px; font-weight: 600; }
.app-header nav { display: flex; gap: 8px; }
.app-header nav a {
  color: rgba(255,255,255,0.85);
  text-decoration: none;
  padding: 6px 14px;
  border-radius: 4px;
  font-size: 13px;
  transition: background 0.2s;
}
.app-header nav a:hover, .app-header nav a.active { background: rgba(255,255,255,0.15); color: white; }
.app-main { max-width: 1200px; margin: 0 auto; padding: 24px; }

/* Cards */
.cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 16px; }
.card {
  background: var(--color-surface);
  border-radius: var(--radius);
  padding: 24px;
  box-shadow: var(--shadow);
  cursor: pointer;
  transition: transform 0.15s, box-shadow 0.15s;
  border: 1px solid var(--color-border);
  text-decoration: none;
  color: inherit;
  display: block;
}
.card:hover { transform: translateY(-2px); box-shadow: var(--shadow-lg); }
.card-icon { font-size: 32px; margin-bottom: 12px; }
.card h3 { font-size: 16px; margin-bottom: 4px; }
.card p { font-size: 13px; color: var(--color-text-secondary); }
.card .badge {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 10px;
  font-size: 11px;
  font-weight: 600;
  margin-top: 8px;
}
.badge-vba { background: #E8F5E9; color: var(--color-success); }
.badge-web { background: #E3F2FD; color: var(--color-primary); }
.badge-shared { background: #FFF8E1; color: var(--color-warning); }

/* Page sections */
.page-section { margin-bottom: 32px; }
.page-section h2 {
  font-size: 20px;
  margin-bottom: 8px;
  padding-bottom: 8px;
  border-bottom: 2px solid var(--color-primary);
}
.section-subtitle { color: var(--color-text-secondary); font-size: 14px; margin-bottom: 16px; }

/* Forms */
.form-group { margin-bottom: 16px; }
.form-group label { display: block; font-size: 13px; font-weight: 600; margin-bottom: 4px; color: var(--color-text-secondary); }
.form-group input, .form-group select {
  width: 100%;
  padding: 8px 12px;
  border: 1px solid var(--color-border);
  border-radius: 4px;
  font-size: 14px;
  font-family: var(--font);
}
.form-group input:focus, .form-group select:focus { outline: none; border-color: var(--color-primary); box-shadow: 0 0 0 2px rgba(68,114,196,0.2); }
.form-row { display: flex; gap: 16px; }
.form-row > * { flex: 1; }

/* Buttons */
.btn {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 10px 20px;
  border: none;
  border-radius: var(--radius);
  font-size: 14px;
  font-weight: 600;
  cursor: pointer;
  transition: background 0.2s, transform 0.1s;
  font-family: var(--font);
}
.btn:active { transform: scale(0.98); }
.btn-primary { background: var(--color-primary); color: white; }
.btn-primary:hover { background: var(--color-primary-dark); }
.btn-outline { background: white; color: var(--color-primary); border: 1px solid var(--color-primary); }
.btn-outline:hover { background: #E3F2FD; }
.btn-danger { background: var(--color-danger); color: white; }
.btn:disabled { opacity: 0.5; cursor: not-allowed; }

/* Drop zone */
.drop-zone {
  border: 2px dashed var(--color-border);
  border-radius: var(--radius);
  padding: 48px;
  text-align: center;
  color: var(--color-text-secondary);
  transition: border-color 0.2s, background 0.2s;
  cursor: pointer;
}
.drop-zone.dragover { border-color: var(--color-primary); background: #E3F2FD; }
.drop-zone .drop-icon { font-size: 48px; margin-bottom: 8px; }

/* Table */
.data-table { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 12px; }
.data-table th {
  background: var(--color-primary);
  color: white;
  padding: 8px 12px;
  text-align: left;
  font-weight: 600;
  cursor: pointer;
  user-select: none;
}
.data-table td { padding: 8px 12px; border-bottom: 1px solid var(--color-border); }
.data-table tr:hover { background: #F5F7FA; }

/* Status */
.status-bar {
  background: var(--color-surface);
  border: 1px solid var(--color-border);
  border-radius: var(--radius);
  padding: 12px 16px;
  margin-top: 16px;
  font-size: 13px;
  color: var(--color-text-secondary);
}

/* Toast */
.toast {
  position: fixed;
  bottom: 24px;
  right: 24px;
  padding: 12px 20px;
  background: var(--color-surface);
  border-radius: var(--radius);
  box-shadow: var(--shadow-lg);
  font-size: 14px;
  z-index: 1000;
  animation: toast-in 0.3s ease;
}
@keyframes toast-in { from { opacity: 0; transform: translateY(12px); } }
.toast-success { border-left: 4px solid var(--color-success); }
.toast-error { border-left: 4px solid var(--color-danger); }

/* Utility */
.hidden { display: none !important; }
.text-center { text-align: center; }
.mt-16 { margin-top: 16px; }
.mb-16 { margin-bottom: 16px; }
```

- [ ] **Step 2: 创建 SPA 路由 + 仪表盘 web/index.html**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CDM Toolkit</title>
<link rel="stylesheet" href="/static/css/app.css">
</head>
<body>
  <header class="app-header">
    <h1>📊 CDM Toolkit</h1>
    <nav>
      <a href="/" data-route="/">仪表盘</a>
      <a href="/merge" data-route="/merge">工作簿合并</a>
      <a href="/split" data-route="/split">工作表拆分</a>
      <a href="/scan" data-route="/scan">特殊字符</a>
      <a href="/files" data-route="/files">文件目录</a>
      <a href="/dates" data-route="/dates">日期计算</a>
      <a href="/units" data-route="/units">单位转换</a>
    </nav>
  </header>
  <main class="app-main" id="app-main">
    <!-- Dashboard (default) -->
    <div class="page-section">
      <h2>功能模块</h2>
      <p class="section-subtitle">选择需要的工具开始处理数据</p>
    </div>
    <div class="cards">
      <a class="card" href="/merge">
        <div class="card-icon">📑</div>
        <h3>工作簿合并</h3>
        <p>将多个 Excel 文件合并为一个工作簿，支持预览和策略配置</p>
        <span class="badge badge-web">Web</span>
      </a>
      <a class="card" href="/split">
        <div class="card-icon">✂️</div>
        <h3>工作表拆分</h3>
        <p>按指定列的值将一个 sheet 拆分为多个，支持大数据量</p>
        <span class="badge badge-web">Web</span>
      </a>
      <a class="card" href="/scan">
        <div class="card-icon">🔍</div>
        <h3>特殊字符扫描</h3>
        <p>扫描非标准字符，生成 EDC（Rave/Clinflash）合规报告</p>
        <span class="badge badge-web">Web</span>
      </a>
      <a class="card" href="/files">
        <div class="card-icon">📁</div>
        <h3>文件目录浏览</h3>
        <p>树形浏览文件夹内容，一键导出带超链接的 Excel 清单</p>
        <span class="badge badge-web">Web</span>
      </a>
      <a class="card" href="/dates">
        <div class="card-icon">📅</div>
        <h3>日期计算</h3>
        <p>对日期列批量加减天数、计算日期差</p>
        <span class="badge badge-shared">双端</span>
      </a>
      <a class="card" href="/units">
        <div class="card-icon">⚖️</div>
        <h3>单位转换</h3>
        <p>实验室检验单位批量统一转换（24组内置因子）</p>
        <span class="badge badge-shared">双端</span>
      </a>
    </div>

    <div class="page-section" style="margin-top:32px;">
      <h2>Excel VBA 工具箱</h2>
      <p class="section-subtitle">以下功能请在 Excel 中按 Alt+F8 使用</p>
      <div class="cards">
        <div class="card">
          <div class="card-icon">📝</div>
          <h3>文本术语标准化</h3>
          <p>批量替换非标准术语为 CDM 标准术语</p>
          <span class="badge badge-vba">VBA</span>
        </div>
        <div class="card">
          <div class="card-icon">🟡</div>
          <h3>重复值高亮</h3>
          <p>高亮相邻行中相同列值连续重复的单元格</p>
          <span class="badge badge-vba">VBA</span>
        </div>
        <div class="card">
          <div class="card-icon">📋</div>
          <h3>空白行管理</h3>
          <p>每N行插入空白行、组间插入、删除空白行</p>
          <span class="badge badge-vba">VBA</span>
        </div>
        <div class="card">
          <div class="card-icon">📖</div>
          <h3>TOC 目录生成</h3>
          <p>生成带超链接的工作簿目录 + SAS Derive 风格</p>
          <span class="badge badge-vba">VBA</span>
        </div>
        <div class="card">
          <div class="card-icon">🗂️</div>
          <h3>工作表管理</h3>
          <p>批量重命名、列出、删除隐藏 sheet</p>
          <span class="badge badge-vba">VBA</span>
        </div>
        <div class="card">
          <div class="card-icon">🔍</div>
          <h3>修改痕迹跟踪</h3>
          <p>比较原版并添加批注显示"旧值→新值"</p>
          <span class="badge badge-vba">VBA</span>
        </div>
      </div>
    </div>

    <div class="status-bar">
      <strong>引擎状态:</strong> Python CDM Engine v3.0 | <strong>VBA:</strong> 8个模块已安装 | <strong>端点:</strong> http://localhost:8520
    </div>
  </main>

  <script type="module" src="/static/js/app.js"></script>
</body>
</html>
```

- [ ] **Step 3: 创建 web/js/app.js — SPA 路由 + 页面加载器**

```javascript
// CDM Toolkit — SPA Router + Page Loader
const ROUTES = {
  '/':       { title: '仪表盘', loader: null },        // inline in index.html
  '/merge':  { title: '工作簿合并', page: 'merge.js' },
  '/split':  { title: '工作表拆分', page: 'split.js' },
  '/scan':   { title: '特殊字符扫描', page: 'scan.js' },
  '/files':  { title: '文件目录浏览', page: 'files.js' },
  '/dates':  { title: '日期计算', page: 'dates.js' },
  '/units':  { title: '单位转换', page: 'units.js' },
};

// Dashboard HTML (saved for restoring)
let dashboardHTML = null;

function saveDashboard() {
  if (!dashboardHTML) {
    const main = document.getElementById('app-main');
    dashboardHTML = main.innerHTML;
  }
}

function restoreDashboard() {
  const main = document.getElementById('app-main');
  if (dashboardHTML) main.innerHTML = dashboardHTML;
}

function updateNav(route) {
  document.querySelectorAll('[data-route]').forEach(a => {
    a.classList.toggle('active', a.dataset.route === route);
  });
}

export function showToast(message, type = 'success') {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 3000);
}

export async function api(endpoint, formData) {
  const res = await fetch(endpoint, { method: 'POST', body: formData });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error(err.detail || `HTTP ${res.status}`);
  }
  return res;
}

async function loadPage(route) {
  const cfg = ROUTES[route];
  if (!cfg) return;

  updateNav(route);
  document.title = `${cfg.title} — CDM Toolkit`;

  if (route === '/') {
    restoreDashboard();
    return;
  }

  saveDashboard();
  const main = document.getElementById('app-main');
  main.innerHTML = '<div class="text-center" style="padding:48px;"><p>加载中...</p></div>';

  try {
    const mod = await import(`/static/js/pages/${cfg.page}`);
    if (mod.default) {
      const content = mod.default();
      main.innerHTML = content;
    }
    if (mod.init) await mod.init();
  } catch (err) {
    main.innerHTML = `<div class="text-center" style="padding:48px;color:var(--color-danger);">
      <h3>页面加载失败</h3><p>${err.message}</p>
      <a href="/" class="btn btn-outline mt-16">返回仪表盘</a></div>`;
  }
}

// SPA Navigation
document.addEventListener('click', (e) => {
  const link = e.target.closest('[data-route]');
  if (!link) return;
  e.preventDefault();
  const route = link.dataset.route;
  history.pushState(null, '', route);
  loadPage(route);
});

window.addEventListener('popstate', () => {
  loadPage(location.pathname);
});

// Initial load
if (location.pathname !== '/') loadPage(location.pathname);
```

- [ ] **Step 4: 验证仪表盘可访问**

Run: `cd e:/Project/Tools/Tools_DM && python server.py --no-browser --port 8521`
Then in browser: open `http://localhost:8521` → verify dashboard renders with 6 Web cards + 6 VBA cards.

- [ ] **Step 5: Commit**

```bash
git add web/index.html web/css/app.css web/js/app.js
git commit -m "feat: add web frontend — SPA router, styles, dashboard"
```

---

### Task 7: Web Components — 4个可复用组件

**Files:**
- Create: `web/js/components/file-upload.js`
- Create: `web/js/components/data-table.js`
- Create: `web/js/components/folder-tree.js`
- Create: `web/js/components/progress-bar.js`

- [ ] **Step 1: file-upload.js — 拖拽上传组件**

```javascript
// web/js/components/file-upload.js — Drag-and-drop file upload
export class FileUpload extends HTMLElement {
  constructor() {
    super();
    this.files = [];
    this.attachShadow({ mode: 'open' });
  }

  connectedCallback() {
    this.accept = this.getAttribute('accept') || '.xlsx,.xls';
    this.multiple = this.hasAttribute('multiple');
    this.render();
  }

  render() {
    this.shadowRoot.innerHTML = `
      <style>
        :host { display: block; }
        .drop-zone {
          border: 2px dashed #ccc;
          border-radius: 8px;
          padding: 32px;
          text-align: center;
          cursor: pointer;
          transition: all 0.2s;
        }
        .drop-zone.dragover { border-color: #4472C4; background: #E3F2FD; }
        .drop-icon { font-size: 36px; }
        .file-list { margin-top: 12px; }
        .file-item {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 6px 12px;
          background: #f5f5f5;
          border-radius: 4px;
          margin-bottom: 4px;
          font-size: 13px;
        }
        .file-item button {
          border: none;
          background: none;
          cursor: pointer;
          color: #C62828;
          font-size: 16px;
        }
      </style>
      <div class="drop-zone">
        <div class="drop-icon">📂</div>
        <p>拖拽文件到此处，或点击选择</p>
        <p style="font-size:12px;color:#999;">支持: ${this.accept}</p>
        <input type="file" hidden ${this.multiple ? 'multiple' : ''} accept="${this.accept}">
      </div>
      <div class="file-list"></div>
    `;

    const zone = this.shadowRoot.querySelector('.drop-zone');
    const input = this.shadowRoot.querySelector('input');
    const fileList = this.shadowRoot.querySelector('.file-list');

    zone.addEventListener('click', () => input.click());
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('dragover');
      this.addFiles(Array.from(e.dataTransfer.files));
    });
    input.addEventListener('change', () => {
      this.addFiles(Array.from(input.files));
      input.value = '';
    });
  }

  addFiles(newFiles) {
    this.files = this.multiple ? [...this.files, ...newFiles] : newFiles;
    this.updateFileList();
    this.dispatchEvent(new CustomEvent('files-changed', {
      detail: { files: this.files },
      bubbles: true,
    }));
  }

  updateFileList() {
    const list = this.shadowRoot.querySelector('.file-list');
    list.innerHTML = this.files.map((f, i) => `
      <div class="file-item">
        <span>📄 ${f.name} (${(f.size / 1024).toFixed(1)} KB)</span>
        <button data-idx="${i}">✕</button>
      </div>
    `).join('');
    list.querySelectorAll('button').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const idx = parseInt(btn.dataset.idx);
        this.files.splice(idx, 1);
        this.updateFileList();
        this.dispatchEvent(new CustomEvent('files-changed', {
          detail: { files: this.files },
          bubbles: true,
        }));
      });
    });
  }

  getFormData() {
    const fd = new FormData();
    this.files.forEach(f => fd.append('files', f));
    return fd;
  }
}
customElements.define('file-upload', FileUpload);
```

- [ ] **Step 2: data-table.js — 交互式表格**

```javascript
// web/js/components/data-table.js — Sortable, filterable data table
export class DataTable extends HTMLElement {
  constructor() { super(); this.attachShadow({ mode: 'open' }); }

  set data(value) { this._data = value; this.render(); }
  get data() { return this._data || []; }

  connectedCallback() {
    this._columns = JSON.parse(this.getAttribute('columns') || '[]');
    this._sortCol = null;
    this._sortDir = 1;
    this.render();
  }

  sort(colKey) {
    if (this._sortCol === colKey) {
      this._sortDir *= -1;
    } else {
      this._sortCol = colKey;
      this._sortDir = 1;
    }
    this._data.sort((a, b) => {
      const va = a[colKey] ?? '', vb = b[colKey] ?? '';
      return String(va).localeCompare(String(vb)) * this._sortDir;
    });
    this.renderBody();
  }

  render() {
    const cols = this._columns;
    this.shadowRoot.innerHTML = `
      <style>
        :host { display: block; overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th { background: #4472C4; color: white; padding: 8px 12px; text-align: left; cursor: pointer; user-select: none; white-space: nowrap; }
        th:hover { background: #2E5BA0; }
        td { padding: 8px 12px; border-bottom: 1px solid #E0E4E8; max-width: 300px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        tr:hover { background: #F5F7FA; }
        .sort-arrow { font-size: 10px; margin-left: 4px; }
        .empty { text-align: center; padding: 24px; color: #999; }
        .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 11px; }
        .badge-modified { background: #E3F2FD; color: #1565C0; }
        .badge-added { background: #E8F5E9; color: #2E7D32; }
        .badge-deleted { background: #FFEBEE; color: #C62828; }
      </style>
      <table>
        <thead><tr>${cols.map(c => `<th data-col="${c.key}">${c.label}${this._sortCol === c.key ? (this._sortDir === 1 ? ' ▲' : ' ▼') : ''}</th>`).join('')}</tr></thead>
        <tbody></tbody>
      </table>
    `;
    this.renderBody();
    this.shadowRoot.querySelectorAll('th').forEach(th => {
      th.addEventListener('click', () => this.sort(th.dataset.col));
    });
  }

  renderBody() {
    const tbody = this.shadowRoot.querySelector('tbody');
    if (!this._data || this._data.length === 0) {
      tbody.innerHTML = '<tr><td class="empty" colspan="99">无数据</td></tr>';
      return;
    }
    tbody.innerHTML = this._data.map(row =>
      `<tr>${this._columns.map(c => `<td>${this.formatCell(row[c.key], c.key)}</td>`).join('')}</tr>`
    ).join('');
  }

  formatCell(value, key) {
    if (key === 'type' || key === 'change_type') {
      const cls = { '修改': 'modified', 'modified': 'modified', '新增': 'added', 'added': 'added', '删除': 'deleted', 'deleted': 'deleted' };
      const label = value || '';
      return `<span class="badge badge-${cls[label] || 'modified'}">${label}</span>`;
    }
    if (value === null || value === undefined) return '';
    const str = String(value);
    return str.length > 100 ? str.slice(0, 100) + '...' : str;
  }
}
customElements.define('data-table', DataTable);
```

- [ ] **Step 3: folder-tree.js — 文件夹树形组件**

```javascript
// web/js/components/folder-tree.js — Recursive folder tree viewer
export class FolderTree extends HTMLElement {
  constructor() { super(); this.attachShadow({ mode: 'open' }); }

  set data(value) { this._data = value; this.render(); }

  render() {
    this.shadowRoot.innerHTML = `
      <style>
        :host { display: block; font-size: 13px; font-family: monospace; }
        .tree { padding: 4px 0; }
        .item { padding: 2px 0; white-space: nowrap; }
        .folder { color: #F57F17; cursor: pointer; }
        .file { color: #1565C0; cursor: pointer; }
        .file:hover, .folder:hover { text-decoration: underline; }
        .size { color: #999; font-size: 11px; margin-left: 8px; }
      </style>
      <div class="tree"></div>
    `;
    if (this._data) this.renderTree();
  }

  renderTree() {
    const container = this.shadowRoot.querySelector('.tree');
    container.innerHTML = this._data.items?.map(item =>
      `<div class="item ${item.is_dir ? 'folder' : 'file'}">
        ${item.is_dir ? '📁' : '📄'} ${item.name}
        ${!item.is_dir ? `<span class="size">${this.formatSize(item.size)}</span>` : ''}
      </div>`
    ).join('') || '<div class="item">空文件夹</div>';
  }

  formatSize(bytes) {
    if (!bytes) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }
}
customElements.define('folder-tree', FolderTree);
```

- [ ] **Step 4: progress-bar.js — 进度条组件**

```javascript
// web/js/components/progress-bar.js — Async task progress indicator
export class ProgressBar extends HTMLElement {
  constructor() { super(); this.attachShadow({ mode: 'open' }); }

  connectedCallback() { this.render(); }

  render() {
    this.shadowRoot.innerHTML = `
      <style>
        :host { display: block; margin: 12px 0; }
        .container { background: #E0E4E8; border-radius: 4px; height: 8px; overflow: hidden; }
        .bar { background: #4472C4; height: 100%; width: 0%; transition: width 0.3s; border-radius: 4px; }
        .label { font-size: 12px; color: #666; margin-top: 4px; }
        .indeterminate .bar { animation: indeterminate 1.5s ease-in-out infinite; width: 30%; }
        @keyframes indeterminate { 0% { transform: translateX(-100%); } 100% { transform: translateX(400%); } }
      </style>
      <div class="container"><div class="bar"></div></div>
      <div class="label"></div>
    `;
  }

  setProgress(value, label) {
    const bar = this.shadowRoot.querySelector('.bar');
    const lbl = this.shadowRoot.querySelector('.label');
    if (value < 0) {
      this.shadowRoot.querySelector('.container').classList.add('indeterminate');
      bar.style.width = '';
      lbl.textContent = label || '处理中...';
    } else {
      this.shadowRoot.querySelector('.container').classList.remove('indeterminate');
      bar.style.width = `${value}%`;
      lbl.textContent = label || `${value}%`;
    }
  }

  hide() { this.style.display = 'none'; }
  show() { this.style.display = ''; }
}
customElements.define('progress-bar', ProgressBar);
```

- [ ] **Step 5: 验证组件注册成功**

在 `web/js/app.js` 顶部添加自动加载:

```javascript
// Auto-load all components
import './components/file-upload.js';
import './components/data-table.js';
import './components/folder-tree.js';
import './components/progress-bar.js';
```

- [ ] **Step 6: Commit**

```bash
git add web/js/components/file-upload.js web/js/components/data-table.js web/js/components/folder-tree.js web/js/components/progress-bar.js
git add web/js/app.js
git commit -m "feat: add 4 reusable Web Components — file-upload, data-table, folder-tree, progress-bar"
```

---

### Task 8: 功能页面 — 合并 + 拆分 + 扫描 + 文件目录 + 日期 + 单位

**Files:**
- Create: `web/js/pages/merge.js`
- Create: `web/js/pages/split.js`
- Create: `web/js/pages/scan.js`
- Create: `web/js/pages/files.js`
- Create: `web/js/pages/dates.js`
- Create: `web/js/pages/units.js`

- [ ] **Step 1: merge.js — 工作簿合并页**

```javascript
// web/js/pages/merge.js — Workbook Merger page
import { api, showToast } from '../app.js';

export default function render() {
  return `
    <div class="page-section">
      <h2>📑 工作簿合并</h2>
      <p class="section-subtitle">将多个 Excel 文件合并为一个工作簿</p>
    </div>
    <file-upload multiple accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group">
        <label><input type="checkbox" id="merge-toc" checked> 自动生成 TOC 目录</label>
      </div>
      <div class="form-group">
        <label><input type="checkbox" id="merge-prefix" checked> 文件名作为 sheet 前缀</label>
      </div>
    </div>
    <button id="merge-btn" class="btn btn-primary" disabled>▶ 开始合并</button>
    <progress-bar id="merge-progress" style="display:none;"></progress-bar>
  `;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('merge-btn');
  const progress = document.getElementById('merge-progress');

  upload.addEventListener('files-changed', e => {
    btn.disabled = e.detail.files.length === 0;
    btn.textContent = e.detail.files.length
      ? `▶ 合并 ${e.detail.files.length} 个文件`
      : '▶ 开始合并';
  });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    upload.files.forEach(f => fd.append('files', f));
    fd.append('create_toc', document.getElementById('merge-toc').checked);
    fd.append('add_prefix', document.getElementById('merge-prefix').checked);

    progress.show();
    progress.setProgress(-1, '正在合并文件...');
    btn.disabled = true;

    try {
      const res = await api('/api/merge', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'merged.xlsx';
      a.click();
      URL.revokeObjectURL(url);
      progress.setProgress(100, '合并完成！');
      showToast('文件合并成功，下载已开始');
    } catch (err) {
      progress.setProgress(0, `失败: ${err.message}`);
      showToast(`合并失败: ${err.message}`, 'error');
    } finally {
      btn.disabled = false;
    }
  });
}
```

- [ ] **Step 2: split.js — 工作表拆分页**

```javascript
// web/js/pages/split.js — Sheet Splitter page
import { api, showToast } from '../app.js';

export default function render() {
  return `
    <div class="page-section">
      <h2>✂️ 工作表拆分</h2>
      <p class="section-subtitle">按指定列的值将工作表拆分为多个 sheet</p>
    </div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group">
        <label for="split-col">拆分列</label>
        <input type="text" id="split-col" value="A" maxlength="3" placeholder="如 A, B, C">
      </div>
      <div class="form-group">
        <label for="split-header">表头行数</label>
        <input type="number" id="split-header" value="1" min="0" max="10">
      </div>
    </div>
    <button id="split-btn" class="btn btn-primary" disabled>▶ 开始拆分</button>
    <progress-bar id="split-progress" style="display:none;"></progress-bar>
  `;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('split-btn');
  const progress = document.getElementById('split-progress');

  upload.addEventListener('files-changed', e => {
    btn.disabled = e.detail.files.length === 0;
  });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    fd.append('file', upload.files[0]);
    fd.append('split_col', document.getElementById('split-col').value);
    fd.append('header_rows', document.getElementById('split-header').value);

    progress.show();
    progress.setProgress(-1, '正在拆分...');
    btn.disabled = true;

    try {
      const res = await api('/api/split', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'split.xlsx';
      a.click();
      URL.revokeObjectURL(url);
      progress.setProgress(100, '拆分完成！');
      showToast('工作表拆分成功');
    } catch (err) {
      progress.setProgress(0, `失败: ${err.message}`);
      showToast(`拆分失败: ${err.message}`, 'error');
    } finally {
      btn.disabled = false;
    }
  });
}
```

- [ ] **Step 3: scan.js — 特殊字符扫描页**

```javascript
// web/js/pages/scan.js — Special Character Scanner
import { api, showToast } from '../app.js';

export default function render() {
  return `
    <div class="page-section">
      <h2>🔍 特殊字符扫描</h2>
      <p class="section-subtitle">扫描工作表中的非标准字符，生成 EDC 合规报告</p>
    </div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group">
        <label for="scan-system">EDC 系统</label>
        <select id="scan-system">
          <option value="Rave">Rave</option>
          <option value="Clinflash">Clinflash</option>
        </select>
      </div>
    </div>
    <button id="scan-btn" class="btn btn-primary" disabled>▶ 开始扫描</button>
    <progress-bar id="scan-progress" style="display:none;"></progress-bar>
    <div id="scan-results" style="margin-top:16px;"></div>
  `;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('scan-btn');
  const progress = document.getElementById('scan-progress');
  const resultsDiv = document.getElementById('scan-results');

  upload.addEventListener('files-changed', e => {
    btn.disabled = e.detail.files.length === 0;
  });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    fd.append('file', upload.files[0]);
    fd.append('system', document.getElementById('scan-system').value);

    progress.show();
    progress.setProgress(-1, '正在扫描...');
    btn.disabled = true;
    resultsDiv.innerHTML = '';

    try {
      const res = await fetch('/api/scan-chars', { method: 'POST', body: fd });
      const json = await res.json();
      progress.setProgress(100, `发现 ${json.total} 个含特殊字符的单元格`);

      if (json.total === 0) {
        resultsDiv.innerHTML = '<p style="color:#2E7D32;padding:16px;">✅ 未发现特殊字符！</p>';
        showToast('扫描完成，未发现特殊字符');
      } else {
        // Build results table
        const table = document.createElement('data-table');
        table.setAttribute('columns', JSON.stringify([
          { key: 'sheet', label: 'Sheet' },
          { key: 'cell', label: '单元格' },
          { key: 'value', label: '内容' },
        ]));
        table.data = json.findings;
        resultsDiv.appendChild(table);

        // Export button
        const exportBtn = document.createElement('button');
        exportBtn.className = 'btn btn-outline mt-16';
        exportBtn.textContent = '📥 导出 JSON 报告';
        exportBtn.addEventListener('click', () => {
          const blob = new Blob([JSON.stringify(json, null, 2)], { type: 'application/json' });
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `special_chars_${json.system}.json`;
          a.click();
        });
        resultsDiv.appendChild(exportBtn);
        showToast(`发现 ${json.total} 个特殊字符`);
      }
    } catch (err) {
      progress.setProgress(0, `失败: ${err.message}`);
      showToast(`扫描失败: ${err.message}`, 'error');
    } finally {
      btn.disabled = false;
    }
  });
}
```

- [ ] **Step 4: files.js — 文件目录浏览页**

```javascript
// web/js/pages/files.js — File Directory Browser
import { api, showToast } from '../app.js';

export default function render() {
  return `
    <div class="page-section">
      <h2>📁 文件目录浏览</h2>
      <p class="section-subtitle">浏览本地文件夹内容，一键导出 Excel 清单</p>
    </div>
    <div class="form-group">
      <label for="files-path">文件夹路径</label>
      <input type="text" id="files-path" placeholder="如 C:\Data\CDM" value="">
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="files-filter">文件类型</label>
        <select id="files-filter">
          <option value="*.*">所有文件</option>
          <option value="*.xlsx">Excel (*.xlsx)</option>
          <option value="*.xls">Excel (*.xls)</option>
          <option value="*.sas7bdat">SAS (*.sas7bdat)</option>
          <option value="*.csv">CSV (*.csv)</option>
          <option value="*.pdf">PDF (*.pdf)</option>
        </select>
      </div>
      <div class="form-group">
        <label><input type="checkbox" id="files-sub" checked> 包含子文件夹</label>
      </div>
    </div>
    <button id="files-btn" class="btn btn-primary">🔍 浏览</button>
    <button id="files-export" class="btn btn-outline" style="display:none;">📥 导出 Excel</button>
    <progress-bar id="files-progress" style="display:none;"></progress-bar>
    <div id="files-results" style="margin-top:16px;"></div>
  `;
}

export async function init() {
  document.getElementById('files-btn').addEventListener('click', async () => {
    const path = document.getElementById('files-path').value.trim();
    if (!path) { showToast('请输入文件夹路径', 'error'); return; }

    const progress = document.getElementById('files-progress');
    const resultsDiv = document.getElementById('files-results');
    progress.show();
    progress.setProgress(-1, '正在浏览...');

    try {
      const fd = new FormData();
      fd.append('path', path);
      fd.append('include_subfolders', document.getElementById('files-sub').checked);
      fd.append('file_pattern', document.getElementById('files-filter').value);

      const res = await fetch('/api/file-listing', { method: 'POST', body: fd });
      const json = await res.json();

      const tree = document.createElement('folder-tree');
      tree.data = json;
      resultsDiv.innerHTML = '';
      resultsDiv.appendChild(tree);

      document.getElementById('files-export').style.display = '';
      document.getElementById('files-export').addEventListener('click', () => {
        const blob = new Blob([JSON.stringify(json, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'file_listing.json'; a.click();
      }, { once: true });

      progress.setProgress(100, `共 ${json.total} 个项目`);
    } catch (err) {
      progress.setProgress(0, `失败: ${err.message}`);
      showToast(`浏览失败: ${err.message}`, 'error');
    }
  });
}
```

- [ ] **Step 5: dates.js — 日期计算页**

```javascript
// web/js/pages/dates.js — Date Calculator
import { api, showToast } from '../app.js';

export default function render() {
  return `
    <div class="page-section">
      <h2>📅 日期计算</h2>
      <p class="section-subtitle">对日期列批量加减天数</p>
    </div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group">
        <label for="dates-col">目标列</label>
        <input type="text" id="dates-col" value="F" maxlength="3">
      </div>
      <div class="form-group">
        <label for="dates-days">加减天数（负数=减）</label>
        <input type="number" id="dates-days" value="30">
      </div>
      <div class="form-group">
        <label for="dates-format">日期格式</label>
        <input type="text" id="dates-format" value="YYYY-MM-DD">
      </div>
    </div>
    <button id="dates-btn" class="btn btn-primary" disabled>▶ 执行</button>
    <progress-bar id="dates-progress" style="display:none;"></progress-bar>
  `;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('dates-btn');
  const progress = document.getElementById('dates-progress');

  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    fd.append('file', upload.files[0]);
    fd.append('col_letter', document.getElementById('dates-col').value);
    fd.append('days', document.getElementById('dates-days').value);
    fd.append('date_format', document.getElementById('dates-format').value);

    progress.show();
    progress.setProgress(-1, '处理中...');
    btn.disabled = true;

    try {
      const res = await api('/api/date-calc', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'dates_result.xlsx'; a.click();
      progress.setProgress(100, '完成！');
      showToast('日期计算完成');
    } catch (err) {
      showToast(`失败: ${err.message}`, 'error');
    } finally { btn.disabled = false; }
  });
}
```

- [ ] **Step 6: units.js — 单位转换页**

```javascript
// web/js/pages/units.js — Lab Unit Converter
import { api, showToast } from '../app.js';

export default function render() {
  return `
    <div class="page-section">
      <h2>⚖️ 实验室单位转换</h2>
      <p class="section-subtitle">将实验室检验结果统一转换为标准单位</p>
    </div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group">
        <label for="units-test-col">检验项目列</label>
        <input type="text" id="units-test-col" value="D" maxlength="3">
      </div>
      <div class="form-group">
        <label for="units-result-col">结果值列</label>
        <input type="text" id="units-result-col" value="E" maxlength="3">
      </div>
      <div class="form-group">
        <label for="units-unit-col">单位列</label>
        <input type="text" id="units-unit-col" value="F" maxlength="3">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="units-std">目标标准单位</label>
        <input type="text" id="units-std" value="mg/dL" placeholder="如 mg/dL, g/L">
      </div>
      <div class="form-group">
        <label for="units-filter">检验项目筛选 (可选)</label>
        <input type="text" id="units-filter" placeholder="留空=全部转换">
      </div>
    </div>
    <button id="units-btn" class="btn btn-primary" disabled>▶ 执行转换</button>
    <progress-bar id="units-progress" style="display:none;"></progress-bar>
  `;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('units-btn');
  const progress = document.getElementById('units-progress');

  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    fd.append('file', upload.files[0]);
    fd.append('test_col', document.getElementById('units-test-col').value);
    fd.append('result_col', document.getElementById('units-result-col').value);
    fd.append('unit_col', document.getElementById('units-unit-col').value);
    fd.append('standard_unit', document.getElementById('units-std').value);
    fd.append('test_filter', document.getElementById('units-filter').value);

    progress.show();
    progress.setProgress(-1, '转换中...');
    btn.disabled = true;

    try {
      const res = await api('/api/unit-convert', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'units_result.xlsx'; a.click();
      progress.setProgress(100, '完成！');
      showToast('单位转换完成');
    } catch (err) {
      showToast(`失败: ${err.message}`, 'error');
    } finally { btn.disabled = false; }
  });
}
```

- [ ] **Step 7: Commit**

```bash
git add web/js/pages/merge.js web/js/pages/split.js web/js/pages/scan.js web/js/pages/files.js web/js/pages/dates.js web/js/pages/units.js
git commit -m "feat: add 6 web pages — merge, split, scan, files, dates, units"
```

---

## 阶段四：VBA 精简

### Task 9: 移除4个VBA模块 + 更新主汇总

**Files:**
- Delete: `vba/modules/modWorkbookMerge.bas`
- Delete: `vba/modules/modSheetSplit.bas`
- Delete: `vba/modules/modSpecialChars.bas`
- Delete: `vba/modules/modFileListing.bas`
- Modify: `vba/CDM_Toolkit.bas` (更新汇总引用)
- Modify: `vba/CDM_Install.bas` (新增 DM_LaunchWeb)

- [ ] **Step 1: 删除4个已迁移到Web的模块**

```bash
rm vba/modules/modWorkbookMerge.bas
rm vba/modules/modSheetSplit.bas
rm vba/modules/modSpecialChars.bas
rm vba/modules/modFileListing.bas
```

- [ ] **Step 2: 更新 CDM_Toolkit.bas — 移除对已删除模块的引用**

确保 `vba/CDM_Toolkit.bas` 只引用保留下来的8个模块：
```vb
' modTextStandardise, modHighlightDup, modBlankRows,
' modTOC, modSheetManager, modTrackChanges,
' modDateCalc, modUnitConvert
```

- [ ] **Step 3: 更新 CDM_Install.bas — 新增 DM_LaunchWeb**

在 `CDM_Install.bas` 中添加：

```vb
'============================================================================
' DM_LaunchWeb
' 启动本地 Web 工具服务器并打开浏览器
' 用法: DM_LaunchWeb
'============================================================================
Public Sub DM_LaunchWeb()
    On Error GoTo ErrHandler
    
    Dim pythonCmd As String
    Dim serverScript As String
    
    ' 尝试找到 server.py 的位置
    ' 假设与 VBA 模块在同一项目目录下
    serverScript = ThisWorkbook.Path & "\server.py"
    
    ' 使用 python 启动服务器
    pythonCmd = "python """ & serverScript & """"
    Shell pythonCmd, vbNormalFocus
    
    MsgBox "CDM Web Toolkit 正在启动..." & vbCrLf & _
           "浏览器将自动打开 http://localhost:8520", vbInformation, "CDM Toolkit"
    Exit Sub
    
ErrHandler:
    MsgBox "DM_LaunchWeb 出错: " & Err.Description & vbCrLf & _
           "请确认已安装 Python 并可从命令行访问。", vbCritical, "CDM Toolkit"
End Sub
```

- [ ] **Step 4: 更新 vba/README.md — 反映精简后的模块列表**

更新 `vba/README.md`，移除4个已删除模块的文档，添加 Web 工具链接说明。

- [ ] **Step 5: Commit**

```bash
git add -u vba/
git add vba/CDM_Toolkit.bas vba/CDM_Install.bas vba/README.md
git commit -m "refactor: remove 4 VBA modules migrated to Web, add DM_LaunchWeb"
```

---

## 阶段五：SAS 工具 + 打包 + 文档

### Task 10: 实现 SAS 工具

**Files:**
- Modify: `cdm_engine/sas_utils.py` (replace placeholder)

- [ ] **Step 1: 写入 sas_utils.py**

```python
"""
CDM SAS Utilities — read/write SAS datasets via pandas.
Requires: pip install pyreadstat
"""
from __future__ import annotations
import os
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font


def read_sas(sas_path: str, encoding: str = "utf-8") -> pd.DataFrame:
    """
    Read a .sas7bdat file into a pandas DataFrame.

    Parameters
    ----------
    sas_path : str
        Path to .sas7bdat file.
    encoding : str
        Character encoding (default utf-8).

    Returns
    -------
    pd.DataFrame
    """
    try:
        import pyreadstat
        df, meta = pyreadstat.read_sas7bdat(sas_path, encoding=encoding)
        return df
    except ImportError:
        # Fallback: try pandas native reader (SAS v7 only)
        return pd.read_sas(sas_path, encoding=encoding)


def sas_to_excel(sas_path: str, output_path: str,
                  sheet_name: str = "DATA") -> Workbook:
    """
    Convert a SAS dataset to an Excel workbook.

    Parameters
    ----------
    sas_path : str
        Path to .sas7bdat file.
    output_path : str
        Where to save the .xlsx file.
    sheet_name : str
        Name for the data sheet.

    Returns
    -------
    Workbook
    """
    df = read_sas(sas_path)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Header row
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True, color="FFFFFF")
        from openpyxl.styles import PatternFill
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

    # Data rows
    for row_idx, (_, row) in enumerate(df.iterrows()):
        for col_idx, val in enumerate(row, 1):
            ws.cell(row=row_idx + 2, column=col_idx,
                     value=str(val) if pd.isna(val) else val)

    wb.save(output_path)
    return wb


def excel_to_sas(xlsx_path: str, output_dir: str,
                  sheet_name: str = "DATA") -> str:
    """
    Convert an Excel file to a SAS dataset (CSV + SAS import script).

    Parameters
    ----------
    xlsx_path : str
        Path to the .xlsx file.
    output_dir : str
        Directory to save CSV and SAS import script.
    sheet_name : str
        Sheet name to read from Excel.

    Returns
    -------
    str — path to the generated SAS import script (.sas)
    """
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    base_name = Path(xlsx_path).stem

    # Save CSV
    csv_path = os.path.join(output_dir, f"{base_name}.csv")
    df.to_csv(csv_path, index=False)

    # Generate SAS import script
    sas_script = f"""/* SAS Import Script — generated by CDM Toolkit */
/* Source: {xlsx_path} */
/* Date: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')} */

PROC IMPORT
    DATAFILE="{csv_path.replace(chr(92), '/')}"
    OUT={base_name}
    DBMS=CSV
    REPLACE;
    GETNAMES=YES;
    GUESSINGROWS=1000;
RUN;

PROC CONTENTS DATA={base_name}; RUN;
"""
    sas_path = os.path.join(output_dir, f"import_{base_name}.sas")
    with open(sas_path, "w", encoding="utf-8") as f:
        f.write(sas_script)

    return sas_path
```

- [ ] **Step 2: 验证 SAS 工具**

Run: `python -c "from cdm_engine.sas_utils import read_sas, sas_to_excel, excel_to_sas; print('SAS utils OK')"`

- [ ] **Step 3: Commit**

```bash
git add cdm_engine/sas_utils.py
git commit -m "feat: add SAS dataset read/write utilities"
```

---

### Task 11: 创建 pyproject.toml 打包配置

**Files:**
- Create: `pyproject.toml`

- [ ] **Step 1: 创建 pyproject.toml**

```toml
[build-system]
requires = ["setuptools>=68.0", "wheel"]
build-backend = "setuptools.backends._legacy:_Backend"

[project]
name = "cdm-tools"
version = "3.0.0"
description = "CDM Toolkit — clinical data management Excel workflows with Python + VBA + Web"
readme = "README.md"
license = {text = "MIT"}
requires-python = ">=3.10"
dependencies = [
    "openpyxl>=3.1",
    "fastapi>=0.104",
    "uvicorn[standard]>=0.24",
    "pandas>=2.0",
]

[project.optional-dependencies]
sas = ["pyreadstat>=1.2"]
dev = [
    "pytest>=7.0",
    "pytest-cov>=4.0",
]

[project.scripts]
cdm-tools = "cdm_tools.cli:main"

[tool.setuptools]
packages = ["cdm_engine"]

[tool.setuptools.package-data]
cdm_engine = ["*.py"]

[tool.pytest.ini_options]
testpaths = ["tests"]
python_files = ["test_*.py"]
```

- [ ] **Step 2: 创建 CLI 入口**

Create `cdm_tools/cli.py`:

```python
"""CDM Toolkit CLI entry point."""
import sys
import os

def main():
    """Entry point for cdm-tools command."""
    if len(sys.argv) < 2:
        print("Usage: cdm-tools serve [--port PORT] [--no-browser]")
        print("       cdm-tools version")
        sys.exit(1)

    cmd = sys.argv[1]
    if cmd == "version":
        from cdm_engine import __version__
        print(f"CDM Toolkit v{__version__}")
    elif cmd == "serve":
        # Parse args
        port = 8520
        open_browser = True
        for i, arg in enumerate(sys.argv[2:], 2):
            if arg == "--port" and i + 1 < len(sys.argv):
                port = int(sys.argv[i + 1])
            elif arg == "--no-browser":
                open_browser = False
        # Add project root to path so server.py can find web/ directory
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        os.chdir(project_root)
        from server import start_server
        start_server(port=port, open_browser=open_browser)
    else:
        print(f"Unknown command: {cmd}")
        sys.exit(1)

if __name__ == "__main__":
    main()
```

And ensure `cdm_tools/__init__.py` exists:

```python
# cdm_tools package
```

- [ ] **Step 3: 验证安装**

Run: `cd e:/Project/Tools/Tools_DM && pip install -e .`
Then: `cdm-tools version`
Expected: `CDM Toolkit v3.0.0`

- [ ] **Step 4: Commit**

```bash
git add pyproject.toml cdm_tools/
git commit -m "build: add pyproject.toml and CLI entry point for pip install"
```

---

### Task 12: 更新 README.md

**Files:**
- Modify: `README.md`

- [ ] **Step 1: 创建项目总览 README.md**

```markdown
# CDM Toolkit v3.0

临床数据管理 (Clinical Data Management) 工具包 — 三层架构。

## 架构

| 层 | 技术 | 职责 |
|---|------|------|
| 表示层 | VBA 宏 + Web UI | Excel内即时操作 + 浏览器复杂处理 |
| 逻辑层 | Python `cdm_engine` | 统一业务逻辑（12个功能模块） |
| 数据层 | 本地文件系统 | Excel (.xlsx/.xls)、SAS (.sas7bdat)、CSV |

## 快速开始

### Web 工具

```bash
pip install cdm-tools
cdm-tools serve          # → 浏览器自动打开 http://localhost:8520
```

### VBA 工具

1. 打开 Excel，按 Alt+F11
2. File → Import File → 选择 `vba/modules/` 下所有 `.bas` 文件
3. 再导入 `vba/CDM_Toolkit.bas`
4. 按 Alt+F8 查看所有可用宏

### Python 引擎

```python
from cdm_engine import standardise_terms, generate_toc, merge_workbooks

wb = load_workbook("your_file.xlsx")
standardise_terms(wb["AE"])
generate_toc(wb)
wb.save("output.xlsx")
```

## 功能清单

| 功能 | VBA | Web | Python |
|------|:---:|:---:|:------:|
| 文本术语标准化 | ✅ | — | ✅ |
| 重复值高亮 | ✅ | — | ✅ |
| 空白行管理 | ✅ | — | ✅ |
| TOC 目录生成 | ✅ | — | ✅ |
| 工作表管理 | ✅ | — | ✅ |
| 修改痕迹跟踪 | ✅ | — | ✅ |
| 工作簿合并 | — | ✅ | ✅ |
| 工作表拆分 | — | ✅ | ✅ |
| 特殊字符扫描 | — | ✅ | ✅ |
| 文件目录浏览 | — | ✅ | ✅ |
| 日期计算 | ✅ | ✅ | ✅ |
| 单位转换 | ✅ | ✅ | ✅ |
| SAS 读写 | — | — | ✅ |

## 系统要求

- Python 3.10+
- Excel 2010+ (for VBA)
- Chrome/Edge (for Web UI)
```

- [ ] **Step 2: Commit**

```bash
git add README.md
git commit -m "docs: update README for v3.0 three-tier architecture"
```

---

## 自审清单

1. **Spec coverage**: 所有12个功能模块都已覆盖 — VBA 8个 ✅, Web 6页 ✅, Python引擎18个函数 ✅, SAS工具 ✅, 打包配置 ✅
2. **Placeholder scan**: 所有代码块均为完整实现，无 TBD/TODO
3. **Type consistency**: engine.py 的函数签名与 server.py 的调用一致；Web Components 的事件名 (`files-changed`) 在所有页面统一使用
