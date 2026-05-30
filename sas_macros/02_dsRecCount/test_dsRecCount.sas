/*****************************************************************************
 *  Test Script: %dsRecCount
 *  Tests dsRecCount by creating two mock datasets in a temporary directory,
 *  running the macro, and printing the result.
 *****************************************************************************/

/* Include the macro definition */
%include "E:\Project\Tools_DM\sas_macros\02_dsRecCount\dsRecCount.sas";

/* Step 1: Create a temporary directory for SAS datasets */
%let tmppath = %sysfunc(pathname(work))\testlib;
%sysexec mkdir "&tmppath";

/* Step 2: Create mock dataset AE (Adverse Events)
   5 rows, 3 variables: USUBJID, AETERM, AESEV */
data work._ae;
    input USUBJID $ AETERM $ AESEV $;
    datalines;
001 Headache MILD
001 Nausea   MODERATE
002 Fatigue  MILD
002 Headache MODERATE
003 Dizziness MILD
;
run;

/* Save to the test directory */
libname _testlib "&tmppath";
data _testlib.ae; set work._ae; run;

/* Step 3: Create mock dataset LB (Lab Results)
   3 rows, 4 variables: USUBJID, LBTEST, LBORRES, LBORRESU */
data work._lb;
    input USUBJID $ LBTEST $ LBORRES $ LBORRESU $;
    datalines;
001 HGB 13.5 g/dL
002 WBC 8.2  10^3/uL
003 PLT 250  10^3/uL
;
run;

data _testlib.lb; set work._lb; run;

/* Step 4: Close the library so dsRecCount opens its own reference */
libname _testlib clear;

/* Step 5: Run dsRecCount on the test directory */
%put === Running dsRecCount on &tmppath ===;
%dsRecCount(lib=&tmppath, label=Test Data Counts);

/* Step 6: Print the result */
proc print data=work.jilu label noobs;
    title "dsRecCount Output";
run;

/* Expected:
   Column e should contain:
     AE  Records:5, Variables:3
     LB  Records:3, Variables:4
*/

/* Step 7: Cleanup — remove test directory */
%sysexec rmdir /s /q "&tmppath";
proc datasets lib=work nolist;
    delete jilu _ae _lb;
quit;
