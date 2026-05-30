/*****************************************************************************
 * Test script for %dropVar
 *
 * Creates 2 datasets in a test source directory:
 *   work.source (via library _src_test):
 *     - dataset_a: USUBJID, __STUDYOID, AGE, SEX
 *     - dataset_b: USUBJID, __STUDYOID, ARM, VISIT
 *
 * Both datasets contain __STUDYOID which will be dropped.
 *
 * Expected output:
 *   - target library contains dataset_a and dataset_b
 *   - __STUDYOID variable removed from both
 *   - All other variables preserved
 *****************************************************************************/

/* Include the macro */
%include "E:/Project/Tools_DM/sas_macros/07_dropVar/dropVar.sas";

title "===== Test: %dropVar =====";

/* Step 1: Create source directory datasets via work library */
/* Use temp folders for portability */
%let srcPath = %sysfunc(getoption(WORK))/dropVar_src;
%let tgtPath = %sysfunc(getoption(WORK))/dropVar_tgt;

/* Remove old folders if they exist */
options noxwait;
x "rmdir /s /q &srcPath 2>NUL";
x "rmdir /s /q &tgtPath 2>NUL";
x "mkdir &srcPath";
x "mkdir &tgtPath";

libname _src "&srcPath";
libname _tgt "&tgtPath";

/* Create dataset_a */
data _src.dataset_a;
    input USUBJID $ __STUDYOID $ AGE SEX $;
    datalines;
SUBJ001 S001 45 M
SUBJ002 S002 38 F
SUBJ003 S003 52 M
;
run;

/* Create dataset_b */
data _src.dataset_b;
    input USUBJID $ __STUDYOID $ ARM $ VISIT $;
    datalines;
SUBJ001 S001 DRUG Week0
SUBJ001 S001 DRUG Week2
SUBJ002 S002 PCB  Week0
;
run;

/* Log input datasets */
title "Input: dataset_a (4 vars: USUBJID, __STUDYOID, AGE, SEX)";
proc print data=_src.dataset_a noobs;
run;

title "Input: dataset_b (4 vars: USUBJID, __STUDYOID, ARM, VISIT)";
proc print data=_src.dataset_b noobs;
run;

/* Step 2: Run dropVar — remove __STUDYOID from all datasets */
%dropVar(var=__STUDYOID, src=&srcPath, tgt=&tgtPath);

/* Step 3: Print output datasets */
title "Output: dataset_a (expected: USUBJID, AGE, SEX — no __STUDYOID)";
proc print data=_tgt.dataset_a noobs;
run;

proc contents data=_tgt.dataset_a;
run;

title "Output: dataset_b (expected: USUBJID, ARM, VISIT — no __STUDYOID)";
proc print data=_tgt.dataset_b noobs;
run;

proc contents data=_tgt.dataset_b;
run;

/* Step 4: Cleanup */
libname _src clear;
libname _tgt clear;

title "===== End Test: %dropVar =====";
