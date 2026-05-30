/*****************************************************************************
 * Test script for %dropVarSuffix
 *
 * Creates 1 dataset in a test source directory with:
 *   - AGE_U   (suffix _U — should be dropped)
 *   - WEIGHT_U (suffix _U — should be dropped)
 *   - HEIGHT   (no suffix — should be retained)
 *   - USUBJID  (no suffix — should be retained)
 *
 * Expected output:
 *   - AGE_U and WEIGHT_U removed
 *   - HEIGHT and USUBJID preserved
 *****************************************************************************/

/* Include the macro */
%include "E:/Project/Tools_DM/sas_macros/08_dropVarSuffix/dropVarSuffix.sas";

title "===== Test: %dropVarSuffix =====";

/* Step 1: Set up test directories */
%let srcPath = %sysfunc(getoption(WORK))/dropVarSuffix_src;
%let tgtPath = %sysfunc(getoption(WORK))/dropVarSuffix_tgt;

options noxwait;
x "rmdir /s /q &srcPath 2>NUL";
x "rmdir /s /q &tgtPath 2>NUL";
x "mkdir &srcPath";
x "mkdir &tgtPath";

libname _src "&srcPath";
libname _tgt "&tgtPath";

/* Create test dataset */
data _src.patient_data;
    input USUBJID $ AGE_U WEIGHT_U HEIGHT;
    datalines;
SUBJ001 45 75.5 170
SUBJ002 38 62.0 165
SUBJ003 52 80.2 178
SUBJ004 29 55.3 160
SUBJ005 41 70.1 172
;
run;

/* Log input dataset */
title "Input: patient_data (4 vars: USUBJID, AGE_U, WEIGHT_U, HEIGHT)";
proc print data=_src.patient_data noobs;
run;

proc contents data=_src.patient_data;
run;

/* Step 2: Run dropVarSuffix — remove all _U variables */
%dropVarSuffix(suffix=_U, src=&srcPath, tgt=&tgtPath);

/* Step 3: Print output dataset */
title "Output: patient_data (expected: USUBJID, HEIGHT only — no AGE_U, WEIGHT_U)";
proc print data=_tgt.patient_data noobs;
run;

proc contents data=_tgt.patient_data;
run;

/* Step 4: Verify the dropped variables are gone */
proc sql;
    select name into :_check separated by ' '
    from dictionary.columns
    where upcase(libname) = "_TGT" and upcase(memname) = "PATIENT_DATA"
      and upcase(name) like "%_U";
quit;

%if %length(&_check) > 0 %then %do;
    %put ERROR: Test FAILED — _U variables still present: &_check;
%end;
%else %do;
    %put NOTE: Test PASSED — all _U variables successfully removed;
%end;

/* Step 5: Cleanup */
libname _src clear;
libname _tgt clear;

title "===== End Test: %dropVarSuffix =====";
