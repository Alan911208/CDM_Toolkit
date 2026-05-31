"""
CDM Toolkit Web Server — FastAPI application.
Start with: cdm-tools serve  (or: python server.py)
"""
import asyncio
import os
import shutil
import subprocess
import sys
import tempfile
import uuid as _uuid
import time as _time
import webbrowser
from pathlib import Path

from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.base import BaseHTTPMiddleware

from cdm_engine import (
    merge_workbooks_from_uploads,
    split_sheet_file,
    scan_special_chars,
    add_days_to_column_file,
    convert_lab_units_file,
    list_directory,
    delete_strikethrough_rows,
    clear_strikethrough,
    create_sheets_from_list,
    delete_sheets,
    calc_cockcroft_gault,
    calc_ckd_epi,
    calc_mdrd_egfr,
    calc_bmi,
    calc_bsa_mosteller,
    calc_ldl_friedewald,
    calc_corrected_calcium,
    calc_ibw_devine,
    calc_anion_gap,
    generate_toc,
    generate_sas_derive_toc,
)

# ── JavaScript MIME fix for Windows ──────────────────────────────────────
# On Windows, Python's mimetypes reads the registry which often maps .js to
# "text/plain".  Browsers *refuse* to load ES modules served that way:
#   "Expected a JavaScript-or-Wasm module script but the server responded
#    with a MIME type of 'text/plain'."
# mimetypes.add_type() alone isn't enough because init() (triggered lazily
# on first guess_type) pulls registry entries that overwrite ours.
# The middleware below fixes Content-Type on the way out — bulletproof.


JS_MIME = "application/javascript"
FIX_EXTENSIONS = (".js", ".mjs")


class JavaScriptMIMEMiddleware(BaseHTTPMiddleware):
    """Ensure .js / .mjs responses always carry a JavaScript Content-Type."""

    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        path = request.url.path.lower()
        if path.endswith(FIX_EXTENSIONS):
            response.headers["Content-Type"] = JS_MIME
        return response


app = FastAPI(title="CDM Toolkit", version="3.0.0")
app.add_middleware(JavaScriptMIMEMiddleware)

WEB_DIR = Path(__file__).parent / "web"
if WEB_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(WEB_DIR)), name="static")


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
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    tmp_dir = tempfile.mkdtemp()
    in_path = os.path.join(tmp_dir, "input.xlsx")
    with open(in_path, "wb") as f:
        f.write(file.file.read())
    wb = load_workbook(in_path)
    all_findings = []
    for ws in wb.worksheets:
        findings = scan_special_chars(ws)
        for row, col, val in findings:
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


# ── Workspace ─────────────────────────────────────────────────────────────
# Session-based workspace for pipeline flow.
# Each browser session gets a temp directory; files persist across requests.
# Session ID from header X-CDM-Session or auto-generated.

_workspaces: dict = {}
_WORKSPACE_TTL = 3600  # 1 hour


def _get_workspace(request: Request) -> dict:
    """Get or create workspace for session."""
    sid = request.headers.get("X-CDM-Session") or request.cookies.get("cdm_session")
    if not sid or sid not in _workspaces:
        sid = _uuid.uuid4().hex[:16]
        ws_dir = tempfile.mkdtemp(prefix="cdm_ws_")
        _workspaces[sid] = {
            "dir": ws_dir,
            "current_file": None,
            "file_name": None,
            "history": [],
            "created_at": _time.time(),
        }
    # Update TTL
    _workspaces[sid]["created_at"] = _time.time()
    return _workspaces[sid]


def _cleanup_workspaces():
    """Remove expired workspaces."""
    now = _time.time()
    expired = [sid for sid, ws in _workspaces.items()
               if now - ws["created_at"] > _WORKSPACE_TTL]
    for sid in expired:
        try:
            shutil.rmtree(_workspaces[sid]["dir"], ignore_errors=True)
        except Exception:
            pass
        del _workspaces[sid]


def _preview_file(file_path: str, rows: int = 100) -> dict:
    """Read first N rows of an Excel file and return as JSON."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    wb = load_workbook(file_path, read_only=True, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        headers = []
        data = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                headers = [str(v) if v is not None else f"Col{j+1}"
                          for j, v in enumerate(row)]
            elif i <= rows:
                data.append([str(v)[:200] if v is not None else "" for v in row])
            else:
                break
        sheets.append({
            "name": ws.title,
            "headers": headers,
            "rows": data,
            "row_count": min(ws.max_row, rows + 1) if ws.max_row else 0,
        })
    wb.close()
    return {"sheets": sheets, "total_sheets": len(sheets)}


@app.post("/api/workspace/upload")
async def ws_upload(request: Request, file: UploadFile = File(...)):
    """Upload a file to the workspace. Returns preview."""
    _cleanup_workspaces()
    ws = _get_workspace(request)
    file_path = os.path.join(ws["dir"], file.filename or "uploaded.xlsx")
    with open(file_path, "wb") as f:
        f.write(file.file.read())
    ws["current_file"] = file_path
    ws["file_name"] = file.filename or "uploaded.xlsx"
    ws["history"].append({"action": "upload", "file": ws["file_name"],
                          "time": _time.strftime("%H:%M:%S")})
    preview = _preview_file(file_path)
    return {
        "session_id": request.headers.get("X-CDM-Session") or "",
        "file_name": ws["file_name"],
        "preview": preview,
        "history": ws["history"][-10:],
    }


@app.get("/api/workspace/preview")
async def ws_preview(request: Request, rows: int = 100):
    """Preview the current workspace file."""
    ws = _get_workspace(request)
    if not ws["current_file"] or not os.path.exists(ws["current_file"]):
        from fastapi.responses import JSONResponse
        return JSONResponse({"error": "工作区无文件"}, status_code=404)
    return _preview_file(ws["current_file"], rows)


@app.get("/api/workspace/info")
async def ws_info(request: Request):
    """Get workspace status."""
    ws = _get_workspace(request)
    info = {
        "has_file": ws["current_file"] is not None,
        "file_name": ws["file_name"],
        "history": ws["history"][-10:],
    }
    if ws["current_file"] and os.path.exists(ws["current_file"]):
        from openpyxl import load_workbook
        try:
            wb = load_workbook(ws["current_file"], read_only=True, data_only=True)
            info["sheets"] = len(wb.sheetnames)
            info["sheet_names"] = wb.sheetnames
            ws_obj = wb.active
            info["rows"] = ws_obj.max_row or 0
            info["cols"] = ws_obj.max_column or 0
            wb.close()
        except Exception:
            info["sheets"] = 0
    return info


@app.post("/api/workspace/apply")
async def ws_apply(request: Request):
    """Apply a tool to the workspace file. Tool-specific params in form data."""
    _cleanup_workspaces()
    ws = _get_workspace(request)
    if not ws["current_file"] or not os.path.exists(ws["current_file"]):
        from fastapi.responses import JSONResponse
        return JSONResponse({"error": "工作区无文件，请先上传"}, status_code=400)

    from openpyxl import load_workbook
    form = await request.form()
    tool = form.get("tool", "")

    wb = load_workbook(ws["current_file"])
    ws_obj = wb.active
    from cdm_engine import (
        generate_toc, generate_sas_derive_toc, standardise_terms,
        highlight_duplicates, insert_blank_every_n, insert_blank_between_groups,
        delete_blank_rows, delete_strikethrough_rows, clear_strikethrough,
        add_days_to_column, convert_lab_units, track_changes,
        create_sheets_from_list, delete_sheets,
        calc_cockcroft_gault, calc_ckd_epi, calc_mdrd_egfr,
        calc_bmi, calc_bsa_mosteller, calc_ldl_friedewald,
        calc_corrected_calcium, calc_ibw_devine, calc_anion_gap,
    )
    from openpyxl.utils import column_index_from_string

    action_detail = tool
    try:
        if tool == "toc":
            style = form.get("style", "standard")
            name = form.get("toc_name", "TOC")
            if style == "derive":
                generate_sas_derive_toc(wb, toc_name=name)
            else:
                generate_toc(wb, toc_name=name)
            action_detail = f"生成{name}目录"

        elif tool == "standardise":
            count = standardise_terms(ws_obj)
            action_detail = f"文本标准化 ({count}处替换)"

        elif tool == "duplicates":
            highlight_duplicates(ws_obj)
            action_detail = "重复值高亮"

        elif tool == "blank-every-n":
            n = int(form.get("n", "5"))
            insert_blank_every_n(ws_obj, n=n)
            action_detail = f"每{n}行插入空白"

        elif tool == "blank-between":
            col = form.get("col", "A")
            insert_blank_between_groups(ws_obj, col_letter=col)
            action_detail = f"按{col}列分组插入空白"

        elif tool == "delete-blank":
            count = delete_blank_rows(ws_obj)
            action_detail = f"删除{count}行空白"

        elif tool == "strikethrough-delete":
            count = delete_strikethrough_rows(ws_obj)
            action_detail = f"删除{count}行删除线"

        elif tool == "strikethrough-clear":
            count = clear_strikethrough(ws_obj)
            action_detail = f"清除{count}个删除线"

        elif tool == "date-add":
            col = form.get("col", "A")
            days = int(form.get("days", "0"))
            add_days_to_column(ws_obj, col, days)
            action_detail = f"日期列{col}{days:+d}天"

        elif tool == "unit-convert":
            tc = form.get("test_col", "A")
            rc = form.get("result_col", "B")
            uc = form.get("unit_col", "C")
            su = form.get("standard_unit", "mg/dL")
            convert_lab_units(ws_obj, tc, rc, uc, su)
            action_detail = f"单位转换为{su}"

        elif tool == "create-sheets":
            names = [n.strip() for n in form.get("sheet_names", "").replace("\n", ",").split(",") if n.strip()]
            create_sheets_from_list(wb, names)
            action_detail = f"创建{len(names)}个工作表"

        elif tool == "delete-sheets":
            names = [n.strip() for n in form.get("sheet_names", "").replace("\n", ",").split(",") if n.strip()]
            delete_sheets(wb, names)
            action_detail = f"删除{len(names)}个工作表"

        elif tool.startswith("medical-"):
            method = tool.split("-", 1)[1]
            row_start = int(form.get("row_start", "2"))
            result_c = column_index_from_string(form.get("result_col", "H"))
            count = 0
            for r in range(row_start, ws_obj.max_row + 1):
                try:
                    def _v(col_key, default=0.0):
                        c = column_index_from_string(form.get(col_key, "A")) if form.get(col_key) else 0
                        v = ws_obj.cell(row=r, column=c).value if c else None
                        return float(v) if v is not None else default

                    def _g(col_key):
                        c = column_index_from_string(form.get(col_key, "A")) if form.get(col_key) else 0
                        g = str(ws_obj.cell(row=r, column=c).value or "").upper() if c else ""
                        return g.startswith("F") or "女" in g

                    if method == "CG":
                        res = calc_cockcroft_gault(_v("age_col"), _v("weight_col"), _v("scr_col"), _g("gender_col"))
                    elif method == "CKD":
                        res = calc_ckd_epi(_v("age_col"), _v("scr_col"), _g("gender_col"))
                    elif method == "MDRD":
                        res = calc_mdrd_egfr(_v("age_col"), _v("scr_col"), _g("gender_col"))
                    elif method == "BMI":
                        res = calc_bmi(_v("weight_col"), _v("height_col"))
                    elif method == "BSA":
                        res = calc_bsa_mosteller(_v("weight_col"), _v("height_col"))
                    elif method == "LDL":
                        res = calc_ldl_friedewald(_v("tc_col"), _v("hdl_col"), _v("tg_col"))
                    elif method == "CORRCA":
                        res = calc_corrected_calcium(_v("ca_col"), _v("albumin_col"))
                    elif method == "IBW":
                        res = calc_ibw_devine(_v("height_col"), _g("gender_col"))
                    elif method == "AG":
                        res = calc_anion_gap(_v("na_col"), _v("cl_col"), _v("hco3_col"))
                    else:
                        continue

                    if result_c > 0 and res != -1.0:
                        ws_obj.cell(row=r, column=result_c).value = res
                        ws_obj.cell(row=r, column=result_c).number_format = '0.00'
                        count += 1
                except (ValueError, TypeError):
                    continue
            action_detail = f"医学计算 ({method}) — {count}行"

        elif tool == "track-changes":
            orig_path = ws["current_file"]  # need original file separately uploaded
            # For workspace, track changes against a copy saved before
            backup = os.path.join(ws["dir"], "_backup.xlsx")
            if os.path.exists(backup):
                track_changes(ws_obj, backup)
                action_detail = "修改痕迹跟踪"
            else:
                return JSONResponse({"error": "需要先保存原始备份"}, status_code=400)

        else:
            return JSONResponse({"error": f"不支持的工具: {tool}"}, status_code=400)

    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

    # Save result back to workspace
    out_path = os.path.join(ws["dir"], "workspace_output.xlsx")
    wb.save(out_path)
    wb.close()
    ws["current_file"] = out_path
    ws["file_name"] = "workspace_output.xlsx"
    ws["history"].append({"action": action_detail, "time": _time.strftime("%H:%M:%S")})

    preview = _preview_file(out_path)
    return {
        "action": action_detail,
        "preview": preview,
        "history": ws["history"][-10:],
    }


@app.get("/api/workspace/download")
async def ws_download(request: Request):
    """Download current workspace file."""
    ws = _get_workspace(request)
    if not ws["current_file"] or not os.path.exists(ws["current_file"]):
        from fastapi.responses import JSONResponse
        return JSONResponse({"error": "工作区无文件"}, status_code=404)
    return FileResponse(ws["current_file"], filename=ws["file_name"] or "output.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── Strikethrough ─────────────────────────────────────────────────────────


@app.post("/api/strikethrough")
async def api_strikethrough(
    file: UploadFile = File(...),
    action: str = Form("delete"),
):
    """Delete strikethrough rows or clear strikethrough formatting."""
    from openpyxl import load_workbook
    in_path = os.path.join(tempfile.mkdtemp(), "input.xlsx")
    with open(in_path, "wb") as f:
        f.write(file.file.read())
    wb = load_workbook(in_path)
    ws = wb.active

    if action == "delete":
        count = delete_strikethrough_rows(ws)
        msg = f"删除了 {count} 行含删除线的数据"
    else:
        count = clear_strikethrough(ws)
        msg = f"清除了 {count} 个单元格的删除线格式"

    out_path = os.path.join(tempfile.mkdtemp(), "output.xlsx")
    wb.save(out_path)
    return FileResponse(out_path, filename="strikethrough_result.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── Sheet Management ──────────────────────────────────────────────────────


@app.post("/api/create-sheets")
async def api_create_sheets(
    file: UploadFile = File(...),
    sheet_names: str = Form(...),
    clear_first: bool = Form(False),
):
    """Create sheets from a comma/newline-separated name list."""
    from openpyxl import load_workbook
    in_path = os.path.join(tempfile.mkdtemp(), "input.xlsx")
    with open(in_path, "wb") as f:
        f.write(file.file.read())
    wb = load_workbook(in_path)

    names = [n.strip() for n in sheet_names.replace("\n", ",").split(",") if n.strip()]
    count = create_sheets_from_list(wb, names, clear_first=clear_first)

    out_path = os.path.join(tempfile.mkdtemp(), "output.xlsx")
    wb.save(out_path)
    return FileResponse(out_path, filename="sheets_result.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/delete-sheets")
async def api_delete_sheets(
    file: UploadFile = File(...),
    sheet_names: str = Form(...),
):
    """Delete specified sheets by name."""
    from openpyxl import load_workbook
    in_path = os.path.join(tempfile.mkdtemp(), "input.xlsx")
    with open(in_path, "wb") as f:
        f.write(file.file.read())
    wb = load_workbook(in_path)

    names = [n.strip() for n in sheet_names.replace("\n", ",").split(",") if n.strip()]
    count = delete_sheets(wb, names)

    out_path = os.path.join(tempfile.mkdtemp(), "output.xlsx")
    wb.save(out_path)
    return FileResponse(out_path, filename="sheets_result.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── Medical Calculators ───────────────────────────────────────────────────


# Column specs for each method: {param_name: (label, required)}
MEDICAL_CALC_METHODS = {
    "CG": {
        "name": "Cockcroft-Gault (CrCl)",
        "desc": "肌酐清除率 — 需要年龄/体重/肌酐/性别",
        "params": {"age_col": True, "scr_col": True, "gender_col": True, "weight_col": True, "result_col": True},
    },
    "CKD": {
        "name": "CKD-EPI 2021 (eGFR)",
        "desc": "估算肾小球滤过率 — 需要年龄/肌酐/性别",
        "params": {"age_col": True, "scr_col": True, "gender_col": True, "result_col": True},
    },
    "MDRD": {
        "name": "MDRD (eGFR)",
        "desc": "MDRD 研究方程 eGFR — 需要年龄/肌酐/性别",
        "params": {"age_col": True, "scr_col": True, "gender_col": True, "result_col": True},
    },
    "BMI": {
        "name": "BMI (体重指数)",
        "desc": "体重(kg) / 身高(m)² — 需要体重/身高",
        "params": {"weight_col": True, "height_col": True, "result_col": True},
    },
    "BSA": {
        "name": "BSA 体表面积 (Mosteller)",
        "desc": "√(身高cm × 体重kg / 3600) — 需要体重/身高",
        "params": {"weight_col": True, "height_col": True, "result_col": True},
    },
    "LDL": {
        "name": "LDL 胆固醇 (Friedewald)",
        "desc": "LDL = TC - HDL - TG/5 — 需要 TC/HDL/TG",
        "params": {"tc_col": True, "hdl_col": True, "tg_col": True, "result_col": True},
    },
    "CORRCA": {
        "name": "校正血钙",
        "desc": "校正 Ca = Ca + 0.8×(4-Albumin) — 需要 Ca/Albumin",
        "params": {"ca_col": True, "albumin_col": True, "result_col": True},
    },
    "IBW": {
        "name": "理想体重 (Devine)",
        "desc": "IBW = 50/45.5 + 2.3×(身高英寸-60) — 需要身高/性别",
        "params": {"height_col": True, "gender_col": True, "result_col": True},
    },
    "AG": {
        "name": "阴离子间隙",
        "desc": "AG = Na - (Cl + HCO₃) — 需要 Na/Cl/HCO3",
        "params": {"na_col": True, "cl_col": True, "hco3_col": True, "result_col": True},
    },
}


@app.get("/api/medical-methods")
async def api_medical_methods():
    """Return available medical calculation methods and their parameters."""
    return {"methods": [
        {"id": k, "name": v["name"], "desc": v["desc"], "params": list(v["params"].keys())}
        for k, v in MEDICAL_CALC_METHODS.items()
    ]}


@app.post("/api/medical-calc")
async def api_medical_calc(
    file: UploadFile = File(...),
    method: str = Form("CG"),
    age_col: str = Form(""),
    scr_col: str = Form(""),
    gender_col: str = Form(""),
    weight_col: str = Form(""),
    height_col: str = Form(""),
    tc_col: str = Form(""),
    hdl_col: str = Form(""),
    tg_col: str = Form(""),
    ca_col: str = Form(""),
    albumin_col: str = Form(""),
    na_col: str = Form(""),
    cl_col: str = Form(""),
    hco3_col: str = Form(""),
    result_col: str = Form(""),
    row_start: int = Form(2),
):
    """Batch clinical calculation on uploaded workbook."""
    from openpyxl import load_workbook
    from openpyxl.utils import column_index_from_string

    if method not in MEDICAL_CALC_METHODS:
        from fastapi.responses import JSONResponse
        return JSONResponse({"error": f"不支持的方法: {method}"}, status_code=400)

    in_path = os.path.join(tempfile.mkdtemp(), "input.xlsx")
    with open(in_path, "wb") as f:
        f.write(file.file.read())
    wb = load_workbook(in_path)
    ws = wb.active

    def _col(letter: str) -> int:
        return column_index_from_string(letter) if letter else 0

    def _val(r: int, c: int, default: float = 0.0) -> float:
        try:
            v = ws.cell(row=r, column=c).value
            return float(v) if v is not None else default
        except (ValueError, TypeError):
            return default

    def _gender(r: int, c: int) -> bool:
        g = str(ws.cell(row=r, column=c).value or "").upper()
        return g.startswith("F") or "女" in g

    age_c, scr_c, gender_c = _col(age_col), _col(scr_col), _col(gender_col)
    weight_c, height_c = _col(weight_col), _col(height_col)
    tc_c, hdl_c, tg_c = _col(tc_col), _col(hdl_col), _col(tg_col)
    ca_c, albumin_c = _col(ca_col), _col(albumin_col)
    na_c, cl_c, hco3_c = _col(na_col), _col(cl_col), _col(hco3_col)
    result_c = _col(result_col)

    count = 0
    for r in range(row_start, ws.max_row + 1):
        try:
            if method == "CG":
                age, scr, wt = _val(r, age_c), _val(r, scr_c), _val(r, weight_c)
                is_f = _gender(r, gender_c)
                result = calc_cockcroft_gault(age, wt, scr, is_f)
            elif method == "CKD":
                result = calc_ckd_epi(_val(r, age_c), _val(r, scr_c), _gender(r, gender_c))
            elif method == "MDRD":
                result = calc_mdrd_egfr(_val(r, age_c), _val(r, scr_c), _gender(r, gender_c))
            elif method == "BMI":
                result = calc_bmi(_val(r, weight_c), _val(r, height_c))
            elif method == "BSA":
                result = calc_bsa_mosteller(_val(r, weight_c), _val(r, height_c))
            elif method == "LDL":
                result = calc_ldl_friedewald(_val(r, tc_c), _val(r, hdl_c), _val(r, tg_c))
            elif method == "CORRCA":
                result = calc_corrected_calcium(_val(r, ca_c), _val(r, albumin_c))
            elif method == "IBW":
                result = calc_ibw_devine(_val(r, height_c), _gender(r, gender_c))
            elif method == "AG":
                result = calc_anion_gap(_val(r, na_c), _val(r, cl_c), _val(r, hco3_c))
            else:
                continue

            if result_c > 0 and result != -1.0:
                ws.cell(row=r, column=result_c).value = result
                ws.cell(row=r, column=result_c).number_format = '0.00'
                count += 1
        except (ValueError, TypeError):
            continue

    out_path = os.path.join(tempfile.mkdtemp(), "output.xlsx")
    wb.save(out_path)
    return FileResponse(out_path, filename=f"{method.lower()}_result.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── TOC Generator ─────────────────────────────────────────────────────────


@app.post("/api/toc")
async def api_toc(
    file: UploadFile = File(...),
    style: str = Form("standard"),
    toc_name: str = Form("TOC"),
):
    """Generate a Table of Contents sheet (standard or SAS Derive style)."""
    from openpyxl import load_workbook
    in_path = os.path.join(tempfile.mkdtemp(), "input.xlsx")
    with open(in_path, "wb") as f:
        f.write(file.file.read())
    wb = load_workbook(in_path)

    if style == "derive":
        generate_sas_derive_toc(wb, toc_name=toc_name)
    else:
        generate_toc(wb, toc_name=toc_name)

    out_path = os.path.join(tempfile.mkdtemp(), "output.xlsx")
    wb.save(out_path)
    return FileResponse(out_path, filename="toc_result.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── Track Changes ─────────────────────────────────────────────────────────


@app.post("/api/track-changes")
async def api_track_changes(
    file: UploadFile = File(...),
    original: UploadFile = File(...),
    add_comments: bool = Form(True),
    highlight: bool = Form(True),
):
    """Compare current file against original, mark changed cells."""
    from openpyxl import load_workbook
    from cdm_engine import track_changes

    tmp_dir = tempfile.mkdtemp()
    in_path = os.path.join(tmp_dir, "current.xlsx")
    orig_path = os.path.join(tmp_dir, "original.xlsx")
    with open(in_path, "wb") as f:
        f.write(file.file.read())
    with open(orig_path, "wb") as f:
        f.write(original.file.read())

    wb = load_workbook(in_path)
    ws = wb.active
    count = track_changes(ws, orig_path, add_comments=add_comments, highlight_changes=highlight)

    out_path = os.path.join(tmp_dir, "output.xlsx")
    wb.save(out_path)
    return FileResponse(out_path, filename="tracked_changes.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── List Sheets ───────────────────────────────────────────────────────────


@app.post("/api/list-sheets")
async def api_list_sheets(
    file: UploadFile = File(...),
):
    """Return all sheet names and visibility status."""
    from openpyxl import load_workbook
    from cdm_engine import list_sheet_names
    in_path = os.path.join(tempfile.mkdtemp(), "input.xlsx")
    with open(in_path, "wb") as f:
        f.write(file.file.read())
    wb = load_workbook(in_path)
    sheets = []
    for name in wb.sheetnames:
        ws = wb[name]
        sheets.append({
            "name": name,
            "visible": ws.sheet_state == "visible",
        })
    return {"count": len(sheets), "sheets": sheets}


# ── SAS Execution ──────────────────────────────────────────────────────────

SAS_EXE = r"C:\Program Files\SASHome\SASFoundation\9.4\sas.exe"
SAS_CONFIGS = {
    "zh": r"C:\Program Files\SASHome\SASFoundation\9.4\nls\zh\SASV9.CFG",
    "u8": r"C:\Program Files\SASHome\SASFoundation\9.4\nls\u8\SASV9.CFG",
}
SAS_MACROS_DIR = Path(__file__).parent / "sas_macros"

# SAS macro metadata — name, description, category
SAS_MACROS = [
    # Category 1: Data Check & Validation
    {"id": "01_varExist", "name": "varExist", "cat": "数据检查与验证", "desc": "检查数据集中是否存在指定变量", "params": "dsn=, var="},
    {"id": "02_dsRecCount", "name": "dsRecCount", "cat": "数据检查与验证", "desc": "列出目录中所有数据集的记录/变量计数", "params": "lib=, label="},
    {"id": "03_dateVarScan", "name": "dateVarScan", "cat": "数据检查与验证", "desc": "扫描库中所有含 DTC/DAT/TIM 的字符变量", "params": "libref="},
    # Category 2: Data Manipulation
    {"id": "04_dateCut", "name": "dateCut", "cat": "数据操作与清理", "desc": "按截止日期过滤临床数据库", "params": "src=, tgt=, cutoff=, visitVar=, otherVar=, keepAll=FORMATS CO, byVar=Usubjid Visit"},
    {"id": "05_textSplit", "name": "textSplit", "cat": "数据操作与清理", "desc": "将长字符变量（>200）拆分为多个列", "params": "dsn=, var="},
    {"id": "06_stripBlank", "name": "stripBlank", "cat": "数据操作与清理", "desc": "从 CDM 数据集中移除空白/占位行", "params": "inlib=, outlib=, insets=, outsets=, idxVar=SEQNUM, keepPS=NO"},
    {"id": "07_dropVar", "name": "dropVar", "cat": "数据操作与清理", "desc": "从库中所有数据集批量删除变量", "params": "var=, src=, tgt="},
    {"id": "08_dropVarSuffix", "name": "dropVarSuffix", "cat": "数据操作与清理", "desc": "批量删除名称以指定后缀结尾的变量", "params": "suffix=_u, src=, tgt="},
    {"id": "09_mergeVar", "name": "mergeVar", "cat": "数据操作与清理", "desc": "从源数据集跨库合并变量", "params": "form=, vars=, key=, src=, tgt="},
    # Category 3: Data I/O
    {"id": "10_sas2csv", "name": "sas2csv", "cat": "数据转换与I/O", "desc": "将目录中所有 SAS 数据集批量转换为 CSV", "params": "src=, tgt="},
    {"id": "11_csv2sas", "name": "csv2sas", "cat": "数据转换与I/O", "desc": "将 CSV 文件导入为 SAS 数据集", "params": "file=, dsn=csv2sas, naming=XLS2SAS, lrecl=5000"},
    {"id": "12_xlsxImport", "name": "xlsxImport", "cat": "数据转换与I/O", "desc": "将 Excel 工作表批量导入为 SAS 数据集", "params": "file=, out="},
    {"id": "13_sas2xlsx", "name": "sas2xlsx", "cat": "数据转换与I/O", "desc": "将 SAS 数据集导出到单个 Excel 工作簿", "params": "lib=, dir=, name=, ext=xlsx"},
    {"id": "15_xls2sas", "name": "xls2sas", "cat": "数据转换与I/O", "desc": "将 Excel 工作表导入 SAS 数据集", "params": "file=, insheet=, outset=, outlib=work, ..."},
    # Category 4: Date & Version
    {"id": "16_stampRunDate", "name": "stampRunDate", "cat": "日期与版本管理", "desc": "为库中所有数据集添加固定运行日期变量", "params": "lib=, date="},
    {"id": "17_runDateDiff", "name": "runDateDiff", "cat": "日期与版本管理", "desc": "新旧版本之间逐行比较单个数据集", "params": "oldlib=, oldset=, newlib=, newset=, outlib=, outset=, keyvar=, exvar="},
    {"id": "18_batchRunDateDiff", "name": "batchRunDateDiff", "cat": "日期与版本管理", "desc": "将 runDateDiff 批量应用于库中所有数据集", "params": "newLib=, oldLib=, outLib=, keyvar=, exvar="},
    {"id": "21_versionDiff", "name": "versionDiff", "cat": "日期与版本管理", "desc": "版本比较：检测新记录/更新/停用记录", "params": "newLib=, oldLib=, outLib=, keyvar=, exclVar=, dropVar="},
    # Category 5: QC
    {"id": "23_qcLabCheck", "name": "qcLabCheck", "cat": "质量控制", "desc": "实验室数据 4 项自动 QC 检查", "params": "lib=, dsn=, out="},
    {"id": "25_dmQueryReport", "name": "dmQueryReport", "cat": "质量控制", "desc": "生成 DMR Query Summary 表（T5.1-T5.3）为 RTF", "params": "qSheet=, form=, outRTF=, ver=3"},
    # Standalone
    {"id": "rawdata_export", "name": "rawdata_export", "cat": "独立脚本", "desc": "审查原始数据并生成内容报告 Excel", "params": ""},
]

# Runtime tracking: run_id → {status, log, start_time}
_sas_jobs: dict = {}


@app.get("/api/sas-list")
async def api_sas_list():
    """Return all available SAS macros with metadata."""
    return {"macros": SAS_MACROS, "configs": list(SAS_CONFIGS.keys())}


@app.get("/api/sas-load/{macro_id}")
async def api_sas_load(macro_id: str):
    """Load a SAS macro's source code and test file."""
    macro = next((m for m in SAS_MACROS if m["id"] == macro_id), None)
    if not macro:
        return {"error": f"未找到宏: {macro_id}"}

    result = {"macro": macro, "source": "", "test": ""}

    # Load macro source
    macro_dir = SAS_MACROS_DIR / macro_id
    if macro_dir.is_dir():
        for f in macro_dir.iterdir():
            if f.suffix == ".sas" and not f.name.startswith("test_"):
                result["source"] = f.read_text(encoding="utf-8", errors="replace")
            elif f.name.startswith("test_"):
                result["test"] = f.read_text(encoding="utf-8", errors="replace")
        # Load README
        readme = macro_dir / "README.md"
        if readme.exists():
            result["readme"] = readme.read_text(encoding="utf-8", errors="replace")
    elif macro_id == "rawdata_export":
        sas_file = SAS_MACROS_DIR / "rawdata_export.sas"
        if sas_file.exists():
            result["source"] = sas_file.read_text(encoding="utf-8", errors="replace")

    return result


@app.post("/api/sas-run")
async def api_sas_run(request: Request):
    """Execute SAS code and return a job ID for polling."""
    data = await request.json()
    sas_code = data.get("code", "")
    config_key = data.get("config", "zh")

    if not sas_code.strip():
        return {"error": "SAS 代码为空"}

    config_path = SAS_CONFIGS.get(config_key, SAS_CONFIGS["zh"])
    if not os.path.exists(SAS_EXE):
        return {"error": f"SAS 可执行文件不存在: {SAS_EXE}"}
    if not os.path.exists(config_path):
        return {"error": f"SAS 配置文件不存在: {config_path}"}

    job_id = _uuid.uuid4().hex[:12]
    _sas_jobs[job_id] = {"status": "running", "log": "", "start_time": _time.time()}

    # Run SAS in background
    asyncio.create_task(_run_sas(job_id, sas_code, config_path))

    return {"job_id": job_id, "status": "running"}


@app.get("/api/sas-status/{job_id}")
async def api_sas_status(job_id: str):
    """Poll SAS job status and get log output."""
    job = _sas_jobs.get(job_id)
    if not job:
        return {"error": "任务不存在"}
    return {
        "status": job["status"],
        "log": job["log"],
        "elapsed": round(_time.time() - job["start_time"], 1),
    }


async def _run_sas(job_id: str, sas_code: str, config_path: str):
    """Run SAS in a subprocess, capture log."""
    job = _sas_jobs.get(job_id)
    tmp_dir = tempfile.mkdtemp(prefix="cdm_sas_")
    sas_file = os.path.join(tmp_dir, "program.sas")
    log_file = os.path.join(tmp_dir, "sas_output.log")
    lst_file = os.path.join(tmp_dir, "sas_output.lst")

    # Include SASAUTOS so macros can %include helpers from their directory
    sas_autos = str(SAS_MACROS_DIR).replace("\\", "/")
    header = f"""
/* CDM Toolkit — SAS Runner */
OPTIONS NOXWAIT NOXSYNC;
OPTIONS SASAUTOS=("{sas_autos}" SASAUTOS);
%let _cdm_tmpdir = {tmp_dir.replace(chr(92), '/')};
"""
    full_code = header + "\n" + sas_code

    try:
        with open(sas_file, "w", encoding="utf-8") as f:
            f.write(full_code)

        cmd = [
            SAS_EXE,
            "-config", config_path,
            "-sysin", sas_file,
            "-log", log_file,
            "-print", lst_file,
            "-nosplash", "-nologo",
            "-nods", "-nonumber", "-nodate",
        ]

        proc = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )

        try:
            stdout, stderr = await asyncio.wait_for(
                proc.communicate(), timeout=300
            )
        except asyncio.TimeoutError:
            proc.kill()
            if job:
                job["status"] = "timeout"
                job["log"] = "⚠️ SAS 执行超时（300 秒）"
            return

        # Read log
        log_text = ""
        if os.path.exists(log_file):
            with open(log_file, "r", encoding="utf-8", errors="replace") as f:
                log_text = f.read()

        # Append stdout/stderr
        if stdout:
            log_text += "\n" + stdout.decode("utf-8", errors="replace")
        if stderr:
            log_text += "\n[STDERR]\n" + stderr.decode("utf-8", errors="replace")

        if job:
            job["status"] = "completed" if proc.returncode == 0 else "error"
            job["log"] = log_text[-50000:]  # last 50k chars
            job["returncode"] = proc.returncode

    except Exception as e:
        if job:
            job["status"] = "error"
            job["log"] = f"执行异常: {str(e)}"


# ── SPA Fallback (must be LAST route) ───────────────────────────────────────


@app.get("/{full_path:path}", response_class=HTMLResponse)
async def spa_fallback(full_path: str = ""):
    """Serve index.html for all unmatched routes (SPA client-side routing)."""
    index_path = WEB_DIR / "index.html"
    if index_path.exists():
        return HTMLResponse(index_path.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>CDM Toolkit</h1><p>Web files not found.</p>")


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
