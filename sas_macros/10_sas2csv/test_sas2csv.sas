/*****************************************************************************
 * Test script for %sas2csv
 *
 * Creates 2 small datasets in a test source directory:
 *   1. demog — USUBJID, AGE, SEX (3 rows)
 *   2. vitals — USUBJID, VISIT, WEIGHT (4 rows)
 *
 * Exports both to CSV files in a separate target directory.
 *
 * Expected output:
 *   - tgt_dir/demog.csv  (3 data rows + header)
 *   - tgt_dir/vitals.csv (4 data rows + header)
 *****************************************************************************/

/* Include the macro */
%include "E:/Project/Tools_DM/sas_macros/10_sas2csv/sas2csv.sas";

title "===== Test: %sas2csv =====";

/* Step 1: Set up test directories */
%let srcPath = %sysfunc(getoption(WORK))/sas2csv_src;
%let tgtPath = %sysfunc(getoption(WORK))/sas2csv_tgt;

options noxwait;
x "rmdir /s /q &srcPath 2>NUL";
x "rmdir /s /q &tgtPath 2>NUL";
x "mkdir &srcPath";
x "mkdir &tgtPath";

libname _src "&srcPath";

/* Create demog dataset — subject demographics */
data _src.demog;
    input USUBJID $ AGE SEX $;
    label
        USUBJID = "Subject ID"
        AGE     = "Age (years)"
        SEX     = "Sex";
    datalines;
SUBJ001 45 M
SUBJ002 38 F
SUBJ003 52 M
;
run;

/* Create vitals dataset — vital signs */
data _src.vitals;
    input USUBJID $ VISIT $ WEIGHT;
    label
        USUBJID = "Subject ID"
        VISIT   = "Visit Name"
        WEIGHT  = "Weight (kg)";
    datalines;
SUBJ001 Week0  75.5
SUBJ001 Week2  74.2
SUBJ002 Week0  62.0
SUBJ002 Week2  61.5
;
run;

/* Log input datasets */
title "Input: demog (3 rows)";
proc print data=_src.demog noobs; run;

title "Input: vitals (4 rows)";
proc print data=_src.vitals noobs; run;

/* Step 2: Run sas2csv — export all datasets to CSV */
%sas2csv(src=&srcPath, tgt=&tgtPath);

/* Step 3: Check that CSV files were created */
data _null_;
    infile "dir /b &tgtPath\*.csv" pipe truncover;
    input fname $100.;
    put "CSV file found: " fname;
run;

/* Step 4: Import demog.csv back to verify content */
proc import datafile="&tgtPath\demog.csv"
    out=work.demog_check
    dbms=csv replace;
    guessingrows=5;
run;

title "Verification: demog imported back from CSV";
proc print data=work.demog_check noobs; run;

/* Step 5: Import vitals.csv back to verify content */
proc import datafile="&tgtPath\vitals.csv"
    out=work.vitals_check
    dbms=csv replace;
    guessingrows=5;
run;

title "Verification: vitals imported back from CSV";
proc print data=work.vitals_check noobs; run;

/* Step 6: Cleanup */
libname _src clear;

title "===== End Test: %sas2csv =====";
