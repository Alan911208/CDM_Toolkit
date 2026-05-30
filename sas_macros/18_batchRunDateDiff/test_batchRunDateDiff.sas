/*****************************************************************************
 * Test Script: batchRunDateDiff
 * Purpose: Verify that %batchRunDateDiff correctly applies runDateDiff to
 *          every dataset in a library.
 *
 * Test Scenario:
 *   - OLD library: 2 datasets (dm, ae) with baseline data and baseline
 *     rundates
 *   - NEW library: same 2 datasets with some changes/additions
 *   - batchRunDateDiff processes both datasets in one call
 *   - Output library: 2 datasets with correct rundate assignments
 *****************************************************************************/

* ---- Step 1: Include the macro (contains _getDatasetList, runDateDiff, batchRunDateDiff) ----;
%include "&_projectRoot\sas_macros\18_batchRunDateDiff\batchRunDateDiff.sas";

* ---- Step 2: Create temporary directories ----;
%let oldDir = %sysfunc(pathname(WORK))\batch_old;
%let newDir = %sysfunc(pathname(WORK))\batch_new;
%let outDir = %sysfunc(pathname(WORK))\batch_out;

%sysexec mkdir "&oldDir";
%sysexec mkdir "&newDir";
%sysexec mkdir "&outDir";

* ---- Step 3: Create mock old datasets ----;

* DM dataset (old) — 3 subjects, baseline date = 01JAN2020;
data dm_old;
    length usubjid $10 siteid $4 armcd $8;
    usubjid = "001-0001"; siteid = "S001"; armcd = "TRT"; rundate = '01JAN2020'd; output;
    usubjid = "001-0002"; siteid = "S001"; armcd = "PBO"; rundate = '01JAN2020'd; output;
    usubjid = "001-0003"; siteid = "S002"; armcd = "TRT"; rundate = '01JAN2020'd; output;
    format rundate date9.;
run;

* AE dataset (old) — 3 events, baseline date = 01JAN2020;
data ae_old;
    length usubjid $10 aeterm $50 aesev $8;
    usubjid = "001-0001"; aeterm = "Headache";  aesev = "MILD";    rundate = '01JAN2020'd; output;
    usubjid = "001-0001"; aeterm = "Nausea";    aesev = "MODERATE"; rundate = '01JAN2020'd; output;
    usubjid = "001-0002"; aeterm = "Fatigue";   aesev = "MILD";    rundate = '01JAN2020'd; output;
    format rundate date9.;
run;

* ---- Step 4: Create mock new datasets (with changes) ----;

* DM dataset (new) — 001-0002 armcd changed, 001-0004 added;
data dm_new;
    length usubjid $10 siteid $4 armcd $8;
    usubjid = "001-0001"; siteid = "S001"; armcd = "TRT"; output;
    usubjid = "001-0002"; siteid = "S001"; armcd = "TRT"; output; * changed from PBO;
    usubjid = "001-0003"; siteid = "S002"; armcd = "TRT"; output;
    usubjid = "001-0004"; siteid = "S003"; armcd = "PBO"; output; * new;
run;

* AE dataset (new) — one severity changed, one event added;
data ae_new;
    length usubjid $10 aeterm $50 aesev $8;
    usubjid = "001-0001"; aeterm = "Headache";  aesev = "MILD";    output;
    usubjid = "001-0001"; aeterm = "Nausea";    aesev = "SEVERE";  output; * severity changed;
    usubjid = "001-0002"; aeterm = "Fatigue";   aesev = "MILD";    output;
    usubjid = "001-0002"; aeterm = "Dizziness"; aesev = "MILD";    output; * new AE;
run;

* ---- Step 5: Copy datasets to directories ----;
libname _old "&oldDir";
libname _new "&newDir";

data _old.dm;  set dm_old;  run;
data _old.ae;  set ae_old;  run;
data _new.dm;  set dm_new;  run;
data _new.ae;  set ae_new;  run;

libname _old clear;
libname _new clear;

* ---- Step 6: Run batchRunDateDiff ----;
%batchRunDateDiff(oldLib=&oldDir, newLib=&newDir, outLib=&outDir,
    keyvar=usubjid aeterm,
    exvar=rundate);

* ---- Step 7: Verify results ----;

libname _out "&outDir";

title "Test 1: DM output — 4 rows, check rundate values";
proc print data=_out.dm noobs;
    format rundate date9.;
run;

title "Test 2: AE output — 4 rows, check rundate values";
proc print data=_out.ae noobs;
    format rundate date9.;
run;

title "Test 3: DM rundate freq — expect 2 old date + 2 today";
proc freq data=_out.dm;
    tables rundate;
    format rundate date9.;
run;

title "Test 4: AE rundate freq — expect 2 old date + 2 today";
proc freq data=_out.ae;
    tables rundate;
    format rundate date9.;
run;

libname _out clear;

* ---- Step 8: Clean up ----;
proc datasets lib=work nolist;
    delete dm_old ae_old dm_new ae_new;
quit;
%sysexec rmdir /s /q "&oldDir";
%sysexec rmdir /s /q "&newDir";
%sysexec rmdir /s /q "&outDir";
title;
