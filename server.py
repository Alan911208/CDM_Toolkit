"""
CDM Toolkit Web Server — FastAPI application.
Start with: cdm-tools serve  (or: python server.py)
"""
import os
import sys
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


@app.get("/{full_path:path}", response_class=HTMLResponse)
async def spa_fallback(full_path: str = ""):
    """Serve index.html for all routes (SPA client-side routing).
    Excludes /api/* and /static/* which are handled by their own routes.
    """
    index_path = WEB_DIR / "index.html"
    if index_path.exists():
        return HTMLResponse(index_path.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>CDM Toolkit</h1><p>Web files not found.</p>")


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
