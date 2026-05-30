/****************************************************************************
 *  Test Script: %qcLabCheck
 *  Purpose: Create a mock lab dataset with deliberate issues and run
 *           qcLabCheck to verify detection of all 4 check types.
 *
 *  Deliberate issues injected:
 *    Check 1 (Missing unit/range):
 *      - Row 2: Glucose result present, LBORRESU missing
 *      - Row 3: ALT result present, LBORNRLO/HI missing
 *    Check 2 (Range outlier >10x):
 *      - Row 5: ALT = 500 (ULN=40) — 12.5x ULN, should flag
 *      - Row 8: WBC = 0.05 (LLN=4.0) — *10 < LLN, should flag
 *    Check 3 (Inconsistent units):
 *      - Row 7: Hemoglobin = "g/dL" for subject 001
 *      - Row 9: Hemoglobin = "mg/dL" for same subject (flag)
 *      - Row 10: Hemoglobin = "g/dL" for subject 003 (also flag)
 *    Check 4 (Non-numeric results):
 *      - Row 11: AST result "<5" — contains "<" character
 *      - Row 12: ALT result ">500" — contains ">" character
 *
 *  Normal records (should NOT be flagged):
 *      - Row 1, 4, 6
 ****************************************************************************/

/* ---- Clean workspace ---- */
proc datasets library=work nolist kill; quit;

/* ---- Create a temporary library path ---- */
%let _testLib = %sysfunc(pathname(work));

/* ---- Build mock LB dataset ---- */
data _lb_raw;
    length usubjid $10 siteid $5 visit $20 form $10
           LBTEST $40 LBORRES $20 LBORRESU $20
           LBORNRLO $20 LBORNRHI $20 LBSIG $20;

    /* Row 1: Normal Glucose — no flags */
    usubjid = "001"; siteid = "01"; visit = "Week 0"; form = "LB";
    LBTEST = "Glucose"; LBORRES = "5.2"; LBORRESU = "mmol/L";
    LBORNRLO = "3.9"; LBORNRHI = "6.1"; LBSIG = ""; output;

    /* Row 2: Missing unit */
    usubjid = "001"; siteid = "01"; visit = "Week 0"; form = "LB";
    LBTEST = "Glucose"; LBORRES = "5.2"; LBORRESU = "";
    LBORNRLO = "3.9"; LBORNRHI = "6.1"; LBSIG = ""; output;

    /* Row 3: Missing normal range */
    usubjid = "002"; siteid = "01"; visit = "Week 0"; form = "LB";
    LBTEST = "ALT"; LBORRES = "25"; LBORRESU = "U/L";
    LBORNRLO = ""; LBORNRHI = ""; LBSIG = ""; output;

    /* Row 4: Normal ALT — no flags */
    usubjid = "002"; siteid = "01"; visit = "Week 4"; form = "LB";
    LBTEST = "ALT"; LBORRES = "28"; LBORRESU = "U/L";
    LBORNRLO = "10"; LBORNRHI = "40"; LBSIG = ""; output;

    /* Row 5: ALT outlier >10x ULN (500 vs 40) */
    usubjid = "002"; siteid = "01"; visit = "Week 8"; form = "LB";
    LBTEST = "ALT"; LBORRES = "500"; LBORRESU = "U/L";
    LBORNRLO = "10"; LBORNRHI = "40"; LBSIG = ""; output;

    /* Row 6: Normal Hemoglobin — no flags (g/dL) */
    usubjid = "001"; siteid = "01"; visit = "Week 0"; form = "LB";
    LBTEST = "Hemoglobin"; LBORRES = "14.5"; LBORRESU = "g/dL";
    LBORNRLO = "12.0"; LBORNRHI = "16.0"; LBSIG = ""; output;

    /* Row 7: Hemoglobin in mg/dL — unit inconsistency */
    usubjid = "001"; siteid = "01"; visit = "Week 4"; form = "LB";
    LBTEST = "Hemoglobin"; LBORRES = "145"; LBORRESU = "mg/dL";
    LBORNRLO = "120"; LBORNRHI = "160"; LBSIG = ""; output;

    /* Row 8: WBC outlier below LLN (*10 < LLN) */
    usubjid = "003"; siteid = "02"; visit = "Week 1"; form = "LB";
    LBTEST = "WBC"; LBORRES = "0.05"; LBORRESU = "10^9/L";
    LBORNRLO = "4.0"; LBORNRHI = "10.0"; LBSIG = ""; output;

    /* Row 9: Hemoglobin again in mg/dL for the same test (flag) */
    usubjid = "003"; siteid = "02"; visit = "Week 1"; form = "LB";
    LBTEST = "Hemoglobin"; LBORRES = "138"; LBORRESU = "mg/dL";
    LBORNRLO = "120"; LBORNRHI = "160"; LBSIG = ""; output;

    /* Row 10: Hemoglobin in g/dL again (unit switch back, flag) */
    usubjid = "003"; siteid = "02"; visit = "Week 4"; form = "LB";
    LBTEST = "Hemoglobin"; LBORRES = "13.8"; LBORRESU = "g/dL";
    LBORNRLO = "12.0"; LBORNRHI = "16.0"; LBSIG = ""; output;

    /* Row 11: Non-numeric result "<5" */
    usubjid = "004"; siteid = "02"; visit = "Week 0"; form = "LB";
    LBTEST = "AST"; LBORRES = "<5"; LBORRESU = "U/L";
    LBORNRLO = "10"; LBORNRHI = "40"; LBSIG = "Abnormal"; output;

    /* Row 12: Non-numeric result ">500" */
    usubjid = "004"; siteid = "02"; visit = "Week 0"; form = "LB";
    LBTEST = "ALT"; LBORRES = ">500"; LBORRESU = "U/L";
    LBORNRLO = "10"; LBORNRHI = "40"; LBSIG = "Abnormal"; output;
run;

/* Copy to the library path so the lib= parm works */
/* The macro creates a libname to &lib, so we use WORK as the library path */
proc sql noprint;
    create table work.lb_test as select * from _lb_raw;
quit;

/* ---- Run %qcLabCheck ---- */
title "Test: %qcLabCheck — 4 QC checks on mock lab data";

%qcLabCheck(
    lib = &_testLib,
    dsn = lb_test,
    out = &_testLib
);

/* ---- Verify results ---- */

title2 "Check 1: Missing unit/range (expect 2 rows: row2 Glucose, row3 ALT)";
proc print data=_fin1 noobs; var usubjid LBTEST LBORRES _check; run;

title2 "Check 2: Range outlier (expect 2 rows: ALT=500, WBC=0.05)";
proc print data=_fin2 noobs; var usubjid LBTEST LBORRES LBORNRHI LBORNRLO _check; run;

title2 "Check 3: Inconsistent units (expect Hemoglobin flagged)";
proc print data=_fin3 noobs; var usubjid LBTEST LBORRESU _check; run;

title2 "Check 4: Non-numeric results (expect 2 rows: AST<5, ALT>500)";
proc print data=_fin4 noobs; var usubjid LBTEST lborres; run;

title "Test complete.";
