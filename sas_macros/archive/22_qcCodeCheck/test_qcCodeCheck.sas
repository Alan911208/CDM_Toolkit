/****************************************************************************
 *  Test Script: %qcCodeCheck
 *  Purpose: Create mock AE and CM coding datasets with deliberate issues
 *           and run qcCodeCheck to verify detection of all 7 check types.
 *
 *  Deliberate issues injected:
 *    AE-1: "Headache" coded to LLT-001 and LLT-002 (same term, diff LLT code)
 *    AE-2: "Nausea" coded to PT-101 and PT-102 (same term, diff PT code)
 *    AE-3: LLT "Pain" coded to LLT-003 and LLT-004 (same LLT, diff code)
 *    AE-4: LLT="Dizziness" and PT="Dizziness" but LLT-005 vs PT-105
 *    CM-1: "Aspirin" coded to DRG-001 and DRG-002 (same drug, diff code)
 *    CM-2: "Ibuprofen" has ATC names "NSAID-Gen" and "NSAID-Specific"
 *    CM-3: DRG-003 + "Hypertension" has ATC4 "C07AB" and "C07AC"
 *
 *  Note: The macro reads pre-loaded _ae and _cm datasets (no Excel file
 *        needed for this test; aeFile/cmFile are omitted).
 ****************************************************************************/

/* ---- Clean workspace ---- */
proc datasets library=work nolist kill; quit;

/* ---- Build mock AE (MedDRA) coding dataset ---- */
/* Column mapping (name-referenced columns for the test):
   _c1_ = project, _c7_ = AE term, _c8_ = LLT, _c9_ = LLT code,
   _c10_ = PT, _c11_ = PT code */
data _ae;
    length _c1_ $20 _c7_ $40 _c8_ $40 _c9_ $40 _c10_ $40 _c11_ $40;
    /* Normal records — should NOT be flagged */
    _c1_ = "TEST001"; _c7_ = "Headache"; _c8_ = "Headache"; _c9_ = "LLT-001";
                     _c10_ = "Headache"; _c11_ = "PT-101"; output;
    _c1_ = "TEST001"; _c7_ = "Nausea"; _c8_ = "Nausea"; _c9_ = "LLT-010";
                     _c10_ = "Nausea"; _c11_ = "PT-101"; output;

    /* AE-1: Same term (Headache), different LLT code */
    _c1_ = "TEST001"; _c7_ = "Headache"; _c8_ = "Headache"; _c9_ = "LLT-002";
                     _c10_ = "Headache"; _c11_ = "PT-101"; output;

    /* AE-2: Same term (Nausea), different PT code */
    _c1_ = "TEST001"; _c7_ = "Nausea"; _c8_ = "Nausea"; _c9_ = "LLT-010";
                     _c10_ = "Nausea"; _c11_ = "PT-102"; output;

    /* AE-3: Same LLT (Pain), different LLT code */
    _c1_ = "TEST001"; _c7_ = "Pain at site"; _c8_ = "Pain"; _c9_ = "LLT-003";
                     _c10_ = "Application site pain"; _c11_ = "PT-201"; output;
    _c1_ = "TEST001"; _c7_ = "Pain general"; _c8_ = "Pain"; _c9_ = "LLT-004";
                     _c10_ = "Pain"; _c11_ = "PT-202"; output;

    /* AE-4: LLT equals PT but codes differ */
    _c1_ = "TEST001"; _c7_ = "Dizziness"; _c8_ = "Dizziness"; _c9_ = "LLT-005";
                     _c10_ = "Dizziness"; _c11_ = "PT-105"; output;

    /* Other project — should be filtered out */
    _c1_ = "OTHER01"; _c7_ = "Fever"; _c8_ = "Pyrexia"; _c9_ = "LLT-099";
                     _c10_ = "Pyrexia"; _c11_ = "PT-999"; output;
run;

/* ---- Build mock CM (WHODD) coding dataset ---- */
/* Column mapping:
   _c1_ = project, _c8_ = drug name, _c9_ = indication,
   _c17_ = drug code, _c18_ = ATC name, _c26_ = ATC4 */
data _cm;
    length _c1_ $20 _c8_ $40 _c9_ $40 _c17_ $40 _c18_ $40 _c26_ $40;
    /* Normal records — should NOT be flagged */
    _c1_ = "TEST001"; _c8_ = "Aspirin"; _c9_ = "Pain";
        _c17_ = "DRG-001"; _c18_ = "Acetylsalicylic acid"; _c26_ = "B01AC06"; output;
    _c1_ = "TEST001"; _c8_ = "Ibuprofen"; _c9_ = "Pain";
        _c17_ = "DRG-010"; _c18_ = "Ibuprofen-Gen"; _c26_ = "M01AE01"; output;

    /* CM-1: Same drug name (Aspirin), different drug code */
    _c1_ = "TEST001"; _c8_ = "Aspirin"; _c9_ = "Pain";
        _c17_ = "DRG-002"; _c18_ = "Acetylsalicylic acid"; _c26_ = "B01AC06"; output;

    /* CM-2: Same drug (Ibuprofen), different preferred name */
    _c1_ = "TEST001"; _c8_ = "Ibuprofen"; _c9_ = "Pain";
        _c17_ = "DRG-010"; _c18_ = "Ibuprofen-Spec"; _c26_ = "M01AE01"; output;

    /* CM-3: Same drug code + indication, different ATC4 */
    _c1_ = "TEST001"; _c8_ = "Metoprolol"; _c9_ = "Hypertension";
        _c17_ = "DRG-003"; _c18_ = "Metoprolol"; _c26_ = "C07AB02"; output;
    _c1_ = "TEST001"; _c8_ = "Metoprolol"; _c9_ = "Hypertension";
        _c17_ = "DRG-003"; _c18_ = "Metoprolol"; _c26_ = "C07AC01"; output;

    /* Other project — filtered out */
    _c1_ = "OTHER01"; _c8_ = "Warfarin"; _c9_ = "DVT";
        _c17_ = "DRG-099"; _c18_ = "Warfarin"; _c26_ = "B01AA03"; output;
run;

/* ---- Run qcCodeCheck (omit aeFile/cmFile — use pre-loaded datasets) ---- */
title "Test: %qcCodeCheck — MedDRA AE + WHODD CM";

%qcCodeCheck(
    project = TEST001,
    out     = %sysfunc(pathname(work))
);

/* ---- Verify results ---- */

title2 "AE-1: Same term, different LLT code (expect 1 row: Headache)";
%if %sysfunc(exist(_a2, data)) %then %do;
    proc print data=_a2 noobs; run;
%end;
%else %put NOTE: _a2 not found — no AE-1 flag;

title2 "AE-2: Same term, different PT code (expect 1 row: Nausea)";
%if %sysfunc(exist(_b2, data)) %then %do;
    proc print data=_b2 noobs; run;
%end;

title2 "AE-3: Same LLT, different LLT code (expect 1 row: Pain)";
%if %sysfunc(exist(_c2, data)) %then %do;
    proc print data=_c2 noobs; run;
%end;

title2 "AE-4: LLT=PT but codes differ (expect 1 row: Dizziness)";
%if %sysfunc(exist(_d2, data)) %then %do;
    proc print data=_d2 noobs; run;
%end;

title2 "CM-1: Same drug, different drug code (expect 1 row: Aspirin)";
%if %sysfunc(exist(_e2, data)) %then %do;
    proc print data=_e2 noobs; run;
%end;

title2 "CM-2: Same drug, different preferred name (expect 1 row: Ibuprofen)";
%if %sysfunc(exist(_f2, data)) %then %do;
    proc print data=_f2 noobs; run;
%end;

title2 "CM-3: Same code+indication, different ATC (expect 1 row: DRG-003)";
%if %sysfunc(exist(_g2, data)) %then %do;
    proc print data=_g2 noobs; run;
%end;

title "Test complete.";
