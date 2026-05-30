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
