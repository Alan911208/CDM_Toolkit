/*****************************************************************************
 * Test Script: stampRunDate
 * Purpose: Verify that %stampRunDate assigns the correct rundate value to
 *          every dataset in a library.
 *****************************************************************************/

* ---- Step 1: Include the macro ----;
%include "&_projectRoot\sas_macros\16_stampRunDate\stampRunDate.sas";

* ---- Step 2: Create a temporary test directory ----;
%let testDir = %sysfunc(pathname(WORK))\test_lib;
%sysexec mkdir "&testDir";

* ---- Step 3: Create 2 test datasets with mock data ----;
* Note: Create them in WORK first, then copy to the test directory;

data dm;
    length usubjid $10 siteid $4 armcd $8;
    usubjid = "001-0001"; siteid = "S001"; armcd = "TRT"; output;
    usubjid = "001-0002"; siteid = "S001"; armcd = "PBO"; output;
    usubjid = "001-0003"; siteid = "S002"; armcd = "TRT"; output;
run;

data ae;
    length usubjid $10 aeterm $50 aesev $8;
    usubjid = "001-0001"; aeterm = "Headache";       aesev = "MILD";   output;
    usubjid = "001-0001"; aeterm = "Nausea";          aesev = "MODERATE"; output;
    usubjid = "001-0002"; aeterm = "Fatigue";         aesev = "MILD";   output;
    usubjid = "001-0003"; aeterm = "Inc. ALT";        aesev = "SEVERE"; output;
run;

* ---- Step 4: Copy datasets to the test directory via libname ----;
libname testlib "&testDir";
proc copy in=work out=testlib;
    select dm ae;
run;

* ---- Step 5: Run stampRunDate with a known date ----;
%stampRunDate(lib=&testDir, date='07JUL2020'd);

* ---- Step 6: Verify results ----;

title "Test 1: DM dataset — all rows should have rundate = 07JUL2020";
proc print data=testlib.dm noobs;
    format rundate date9.;
run;

title "Test 2: AE dataset — all rows should have rundate = 07JUL2020";
proc print data=testlib.ae noobs;
    format rundate date9.;
run;

title "Test 3: Frequency of rundate in DM (expect 3 rows, all 07JUL2020)";
proc freq data=testlib.dm;
    tables rundate;
    format rundate date9.;
run;

title "Test 4: Frequency of rundate in AE (expect 4 rows, all 07JUL2020)";
proc freq data=testlib.ae;
    tables rundate;
    format rundate date9.;
run;

* ---- Step 7: Clean up ----;
libname testlib clear;
%sysexec rmdir /s /q "&testDir";
title;
