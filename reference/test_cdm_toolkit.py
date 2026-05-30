"""
Test script for CDM Excel Toolkit.
Creates mock clinical data workbooks, runs every toolkit function,
and writes results to test_output/ for manual inspection.
"""

import os, sys, shutil, tempfile
from datetime import datetime, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

# Ensure we can import the toolkit
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cdm_toolkit import *

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_output")


def setup_output_dir():
    """Clean and recreate test_output."""
    if os.path.exists(OUTPUT_DIR):
        shutil.rmtree(OUTPUT_DIR)
    os.makedirs(OUTPUT_DIR)


# ================================================================
#  Mock Data Generators
# ================================================================

def make_ae_data(wb: Workbook):
    """Create a realistic AE (Adverse Events) sheet with deliberate issues."""
    ws = wb.active
    ws.title = "AE"

    headers = ["STUDYID", "SITEID", "USUBJID", "AESEQ", "AETERM",
               "AESTDTC", "AEENDTC", "AESER", "AESEV", "AEOUT", "AEREL"]
    ws.append(headers)
    apply_header_style(ws)

    # 20 rows of AE data with some duplicates and non-standard terms
    subjects = ["001-001", "001-001", "001-002", "001-002", "001-003",
                "001-003", "001-004", "001-004", "001-005", "001-005",
                "001-001", "001-002", "001-003", "001-004", "001-005",
                "001-001", "001-002", "001-003", "001-004", "001-005"]
    terms = ["Headache", "Headache", "Nausea", "Nausea", "Fatigue",
             "Fatigue", "Dizziness", "Dizziness", "Rash", "Rash",
             "Headache", "Nausea", "Fatigue", "Dizziness", "Rash",
             "Headache", "Nausea", "Fatigue", "Dizziness", "Rash"]
    outcomes = ["Recovered", "Ongoing", "Unknown",
                "Recovered/Resolved with Sequelae",
                "Recovered", "Ongoing", "Unknown",
                "Recovered", "Ongoing", "Unknown",
                "Recovered", "Ongoing", "Unknown", "Recovered", "Ongoing",
                "Recovered", "Ongoing", "Unknown", "Recovered", "Ongoing"]
    relationships = ["Reasonable Possib", "Reasonable Possib",
                     "Reasonable Possib", "Reasonable Possib",
                     "Reasonable Possib", "Reasonable Possib",
                     "Reasonable Possib", "Reasonable Possib",
                     "Reasonable Possib", "Reasonable Possib",
                     "Reasonable Possib", "Reasonable Possib",
                     "Reasonable Possib", "Reasonable Possib",
                     "Reasonable Possib", "Reasonable Possib",
                     "Reasonable Possib", "Reasonable Possib",
                     "Reasonable Possib", "Reasonable Possib"]

    start_date = datetime(2025, 6, 1)
    for i in range(20):
        ws.append([
            "STUDY001", f"SITE{10 + i % 3:02d}", subjects[i], i + 1, terms[i],
            start_date + timedelta(days=i * 7),       # datetime, not string
            start_date + timedelta(days=i * 7 + 5),
            "N", "MILD", outcomes[i], relationships[i]
        ])
    auto_width(ws)
    return ws


def make_lab_data(wb: Workbook):
    """Create a lab data sheet with mixed units for testing unit conversion."""
    ws = wb.create_sheet("LB")

    headers = ["STUDYID", "USUBJID", "VISIT", "LBTEST", "LBORRES",
               "LBORRESU", "LBORNRLO", "LBORNRHI", "LBSIG"]
    ws.append(headers)
    apply_header_style(ws)

    # Mix of units: some in mg, some in g, some in mL, some in L
    lab_data = [
        # Glucose (mg/dL and g/L mixed)
        ["STUDY001", "001-001", "Screening", "Glucose", 90,  "mg/dL",  70,  110,  "Normal"],
        ["STUDY001", "001-001", "Week 4",    "Glucose", 0.9,  "g/L",   0.7, 1.1,  "Normal"],
        ["STUDY001", "001-002", "Screening", "Glucose", 85,  "mg/dL",  70,  110,  "Normal"],
        ["STUDY001", "001-002", "Week 4",    "Glucose", 95,  "mg/dL",  70,  110,  "Normal"],
        # Sodium (mmol/L — no conversion needed)
        ["STUDY001", "001-001", "Screening", "Sodium",  140, "mmol/L", 135, 145,  "Normal"],
        ["STUDY001", "001-001", "Week 4",    "Sodium",  138, "mmol/L", 135, 145,  "Normal"],
        # Hemoglobin (g/L and g/dL mixed)
        ["STUDY001", "001-003", "Screening", "Hemoglobin", 14.5, "g/dL", 12, 16, "Normal"],
        ["STUDY001", "001-003", "Week 4",    "Hemoglobin", 145,  "g/L",  120, 160, "Normal"],
        ["STUDY001", "001-004", "Screening", "Hemoglobin", 13.2, "g/dL", 12, 16, "Normal"],
        # Creatinine
        ["STUDY001", "001-001", "Screening", "Creatinine", 0.9,  "mg/dL", 0.6, 1.2, "Normal"],
        ["STUDY001", "001-002", "Screening", "Creatinine", 1.1,  "mg/dL", 0.6, 1.2, "Normal"],
        ["STUDY001", "001-003", "Screening", "Creatinine", 88,   "umol/L", 53, 106, "Normal"],
        # Outlier
        ["STUDY001", "001-005", "Screening", "Glucose", 500,  "mg/dL",  70,  110,  "Abnormal"],
        ["STUDY001", "001-005", "Week 4",    "Glucose", 450,  "mg/dL",  70,  110,  "Abnormal"],
        # Non-numeric result
        ["STUDY001", "001-004", "Screening", "Glucose", ">500", "mg/dL", 70, 110, "Abnormal"],
    ]
    for row in lab_data:
        ws.append(row)
    auto_width(ws)
    return ws


def make_multiple_workbooks():
    """Create 3 small .xlsx files in test_output/sources/ for merge testing."""
    src_dir = os.path.join(OUTPUT_DIR, "sources")
    os.makedirs(src_dir, exist_ok=True)

    for site in ["SITE01", "SITE02", "SITE03"]:
        wb = Workbook()
        ws = wb.active
        ws.title = f"DM_{site}"
        ws.append(["USUBJID", "SITE", "VISIT", "FORM", "VAR", "VALUE"])
        for i in range(1, 6):
            ws.append([f"{site}-00{i}", site, f"Visit {i}", "DM", "BRTHDTC",
                       f"19{80+i}-01-01"])
        wb.save(os.path.join(src_dir, f"{site}_DM.xlsx"))
        wb.close()


# ================================================================
#  Test Runner
# ================================================================

def test_01_text_standardisation():
    print("[TEST 01] Text standardisation ...", end=" ")
    wb = Workbook()
    make_ae_data(wb)
    ws = wb["AE"]

    count = standardise_terms(ws)
    # Verify replaced terms
    sample = ws.cell(row=2, column=10).value  # AEOUT for row 2 was "Recovered"
    assert "Recovered / Resolved" in str(sample), f"Got: {sample}"
    wb.save(os.path.join(OUTPUT_DIR, "test_01_standardised.xlsx"))
    print(f"OK ({count} replacements)")


def test_02_highlight_duplicates():
    print("[TEST 02] Duplicate highlighting ...", end=" ")
    wb = Workbook()
    make_ae_data(wb)
    ws = wb["AE"]

    highlight_duplicates(ws, col_start=5, col_end=5, row_start=2)  # AETERM column
    # Rows 9 and 10 are "Rash","Rash" — both should be yellow
    assert ws.cell(row=9, column=5).fill.start_color.rgb == "00FFFF00" \
        or ws.cell(row=9, column=5).fill.start_color.rgb == "FFFF00"
    wb.save(os.path.join(OUTPUT_DIR, "test_02_duplicates.xlsx"))
    print("OK")


def test_03_blank_row_insertion():
    print("[TEST 03] Blank row insertion ...", end=" ")
    wb = Workbook()
    make_ae_data(wb)
    ws = wb["AE"]

    # Every 5 rows
    inserted = insert_blank_every_n(ws, n=5, col_letter="A")
    assert inserted > 0
    # Between groups (by USUBJID in column C)
    inserted2 = insert_blank_between_groups(ws, col_letter="C")
    assert inserted2 > 0
    wb.save(os.path.join(OUTPUT_DIR, "test_03_blank_rows.xlsx"))
    print(f"OK (every-5: {inserted}, between-groups: {inserted2})")


def test_04_toc_generation():
    print("[TEST 04] TOC generation ...", end=" ")
    wb = Workbook()
    make_ae_data(wb)
    make_lab_data(wb)

    generate_toc(wb, "TOC")
    assert "TOC" in wb.sheetnames
    toc = wb["TOC"]
    # Should have links for AE and LB
    assert toc.cell(row=2, column=1).value in ("AE", "LB")
    wb.save(os.path.join(OUTPUT_DIR, "test_04_toc.xlsx"))
    print("OK")


def test_05_sheet_splitter():
    print("[TEST 05] Sheet splitter ...", end=" ")
    wb = Workbook()
    make_ae_data(wb)
    ws = wb["AE"]

    # Add a column we can split by (column L = severity)
    ws.cell(row=1, column=12, value="SEVERITY_GROUP")
    for i in range(2, min(ws.max_row + 1, 12)):
        ws.cell(row=i, column=12, value="MILD" if i % 3 == 0 else "MODERATE")
    ws.cell(row=2, column=12, value="SEVERE")  # one SEVERE row

    sheets = split_sheet_by_column(wb, "AE", split_col="L", header_rows=1)
    assert "MILD" in wb.sheetnames or "MODERATE" in wb.sheetnames
    wb.save(os.path.join(OUTPUT_DIR, "test_05_split.xlsx"))
    print(f"OK ({len(sheets)} sheets: {sheets})")


def test_06_workbook_merger():
    print("[TEST 06] Workbook merger ...", end=" ")
    make_multiple_workbooks()
    src_dir = os.path.join(OUTPUT_DIR, "sources")
    files = [os.path.join(src_dir, f) for f in os.listdir(src_dir) if f.endswith(".xlsx")]

    out_path = os.path.join(OUTPUT_DIR, "test_06_merged.xlsx")
    merged = merge_workbooks(files, out_path, create_toc=True)
    # Should have TOC + 3 site sheets
    assert "TOC" in merged.sheetnames
    assert len(merged.sheetnames) >= 4
    print(f"OK ({len(merged.sheetnames)} sheets)")


def test_07_sas_derive_toc():
    print("[TEST 07] SAS-Derive-style TOC ...", end=" ")
    wb = Workbook()
    make_ae_data(wb)
    make_lab_data(wb)
    # Write a description in V2 of each sheet
    wb["AE"]["V2"] = "Adverse Events dataset — all AEs recorded"
    wb["LB"]["V2"] = "Laboratory data — chemistry panel"

    generate_sas_derive_toc(wb, desc_col="V2", toc_name="TOC")
    toc = wb["TOC"]
    assert "Adverse Events" in str(toc.cell(row=2, column=2).value)
    wb.save(os.path.join(OUTPUT_DIR, "test_07_sas_derive_toc.xlsx"))
    print("OK")


def test_08_track_changes():
    print("[TEST 08] Track changes ...", end=" ")
    # Save original
    orig_path = os.path.join(OUTPUT_DIR, "test_08_original.xlsx")
    wb = Workbook()
    make_ae_data(wb)
    wb.save(orig_path)

    # Modify
    wb["AE"].cell(row=2, column=5).value = "Migraine"  # was "Headache"
    wb["AE"].cell(row=3, column=10).value = "Resolved"  # was "Ongoing"

    changes = track_changes_add_comment(wb["AE"], orig_path)
    assert changes >= 2
    wb.save(os.path.join(OUTPUT_DIR, "test_08_tracked.xlsx"))
    print(f"OK ({changes} changes flagged)")


def test_09_date_calculator():
    print("[TEST 09] Date calculator ...", end=" ")
    wb = Workbook()
    make_ae_data(wb)
    ws = wb["AE"]

    # Add 30 days to AE start dates (column F)
    n = add_days_to_column(ws, "F", 30, row_start=2)
    assert n > 0

    # Calculate duration: G - F → column M
    ws.cell(row=1, column=13, value="AEDUR_DAYS")
    m = calc_date_diff(ws, "F", "G", "M", row_start=2)
    assert m > 0
    wb.save(os.path.join(OUTPUT_DIR, "test_09_dates.xlsx"))
    print(f"OK (dates shifted: {n}, diffs computed: {m})")


def test_10_lab_unit_conversion():
    print("[TEST 10] Lab unit conversion ...", end=" ")
    wb = Workbook()
    make_lab_data(wb)
    ws = wb["LB"]

    # Convert all Glucose to mg/dL
    n1 = convert_lab_units(ws, "D", "E", "F", "mg/dL", test_filter="Glucose")
    # Convert all Hemoglobin to g/dL
    n2 = convert_lab_units(ws, "D", "E", "F", "g/dL", test_filter="Hemoglobin")

    # Verify Glucose in g/L was converted
    # Row 3 (1-indexed) had Glucose 0.9 g/L → should now be ~90 mg/dL
    for row_idx in range(2, ws.max_row + 1):
        test = ws.cell(row=row_idx, column=4).value
        unit = ws.cell(row=row_idx, column=6).value
        if test == "Glucose":
            assert unit == "mg/dL", f"Row {row_idx}: Glucose unit not mg/dL, got {unit}"

    wb.save(os.path.join(OUTPUT_DIR, "test_10_lab_converted.xlsx"))
    print(f"OK (converted: {n1+n2})")


def test_11_special_char_scanner():
    print("[TEST 11] Special char scanner ...", end=" ")
    wb = Workbook()
    ws = wb.active
    ws.title = "DATA"
    ws.append(["ID", "TEXT"])
    ws.append([1, "Normal ASCII text"])
    ws.append([2, "Contains emoji: 😀 and Chinese: 中文"])
    ws.append([3, "Another normal row"])

    findings = scan_special_chars(ws)
    assert len(findings) >= 1
    wb.save(os.path.join(OUTPUT_DIR, "test_11_special_chars.xlsx"))
    print(f"OK ({len(findings)} cells with special chars)")


def test_12_delete_blank_rows():
    print("[TEST 12] Delete blank rows ...", end=" ")
    wb = Workbook()
    make_ae_data(wb)
    ws = wb["AE"]

    # Insert some blank rows manually
    ws.insert_rows(5)
    ws.insert_rows(10)
    ws.cell(row=5, column=1, value=None)
    ws.cell(row=10, column=1, value=None)

    deleted = delete_blank_rows(ws)
    assert deleted >= 2
    wb.save(os.path.join(OUTPUT_DIR, "test_12_no_blanks.xlsx"))
    print(f"OK ({deleted} blank rows removed)")


def test_13_batch_rename():
    print("[TEST 13] Batch rename sheets ...", end=" ")
    wb = Workbook()
    make_ae_data(wb)
    make_lab_data(wb)

    batch_rename_sheets(wb, {"AE": "Adverse_Events", "LB": "Laboratory"})
    assert "Adverse_Events" in wb.sheetnames
    assert "Laboratory" in wb.sheetnames
    assert "AE" not in wb.sheetnames
    wb.save(os.path.join(OUTPUT_DIR, "test_13_renamed.xlsx"))
    print("OK")


def test_14_full_pipeline():
    """End-to-end test: all steps on a single workbook."""
    print("[TEST 14] Full pipeline ...", end=" ")
    wb = Workbook()
    make_ae_data(wb)
    make_lab_data(wb)

    # 1. Standardise terms
    n = standardise_terms(wb["AE"])
    print(f"\n       Terms standardised: {n}", end="")

    # 2. Highlight duplicates in AETERM
    highlight_duplicates(wb["AE"], col_start=5, col_end=5, row_start=2)

    # 3. Convert lab units
    n2 = convert_lab_units(wb["LB"], "D", "E", "F", "mg/dL", test_filter="Glucose")
    print(f", Lab conversions: {n2}", end="")

    # 4. Add date difference
    wb["AE"].cell(row=1, column=13, value="AEDUR")
    calc_date_diff(wb["AE"], "F", "G", "M", row_start=2)

    # 5. Generate TOC
    generate_toc(wb, "TOC")

    # 6. Auto-width all sheets
    for ws in wb.worksheets:
        auto_width(ws)

    wb.save(os.path.join(OUTPUT_DIR, "test_14_full_pipeline.xlsx"))
    print("\n       Saved test_14_full_pipeline.xlsx — OK")


# ================================================================
#  Main
# ================================================================
if __name__ == "__main__":
    setup_output_dir()
    print(f"CDM Toolkit — Test Suite")
    print(f"Output directory: {OUTPUT_DIR}")
    print("-" * 50)

    tests = [
        test_01_text_standardisation,
        test_02_highlight_duplicates,
        test_03_blank_row_insertion,
        test_04_toc_generation,
        test_05_sheet_splitter,
        test_06_workbook_merger,
        test_07_sas_derive_toc,
        test_08_track_changes,
        test_09_date_calculator,
        test_10_lab_unit_conversion,
        test_11_special_char_scanner,
        test_12_delete_blank_rows,
        test_13_batch_rename,
        test_14_full_pipeline,
    ]

    passed = 0
    failed = 0
    for test in tests:
        try:
            test()
            passed += 1
        except Exception as e:
            print(f"FAIL — {e}")
            failed += 1
            import traceback; traceback.print_exc()

    print("-" * 50)
    print(f"Results: {passed} PASSED, {failed} FAILED out of {len(tests)} tests")
    print(f"Output files: {OUTPUT_DIR}")
