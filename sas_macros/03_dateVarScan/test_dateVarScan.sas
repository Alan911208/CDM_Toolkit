/*****************************************************************************
 *  Test Script: %dateVarScan
 *  Creates two mock datasets — one with DTC variables (AE), one without (DM).
 *  Places them in a temp library, runs dateVarScan, and prints results.
 *****************************************************************************/

/* Include the macro definition (includes internal helper _dropIfExists) */
%include "E:\Project\Tools_DM\sas_macros\03_dateVarScan\dateVarScan.sas";

/* Step 1: Create a temporary directory for the test library */
%let tmppath = %sysfunc(pathname(work))\datetest;
%sysexec mkdir "&tmppath";

/* Step 2: Create mock AE dataset with date/time variables
   Variables: STUDYID, USUBJID, AESTDTC, AEENDTC, AETERM */
data work._ae;
    length STUDYID $20 USUBJID $10 AESTDTC $16 AEENDTC $16 AETERM $40;
    input STUDYID $ USUBJID $ AESTDTC $ AEENDTC $ AETERM $ &;
    datalines;
STUDY001 001 2024-01-15T08:00 2024-01-16T12:00 Headache
STUDY001 001 2024-02-20T09:30 2024-02-22T14:00 Nausea
STUDY001 002 2024-03-10T10:00 2024-03-10T18:00 Fatigue
STUDY001 003 2024-04-05T07:15             Dizziness
;
run;

/* Step 3: Create mock DM dataset WITHOUT date/time variables
   Variables: STUDYID, USUBJID, SEX, AGE, ARM */
data work._dm;
    input STUDYID $ USUBJID $ SEX $ AGE ARM $;
    datalines;
STUDY001 001 M 45 TRT
STUDY001 002 F 52 CTL
STUDY001 003 F 38 TRT
;
run;

/* Step 4: Save datasets to the test library */
libname _dlib "&tmppath";
data _dlib.ae; set work._ae; run;
data _dlib.dm; set work._dm; run;

/* Step 5: Run dateVarScan */
%put === Running dateVarScan on &tmppath ===;
%dateVarScan(libref=_dlib);

/* Step 6: Print the date variable inventory */
%if %sysfunc(exist(_dlib.datelist, data)) %then %do;
    proc print data=_dlib.datelist noobs;
        title "dateVarScan Output — Date/Time Variable Inventory";
    run;

    /* Summarize findings */
    proc freq data=_dlib.datelist;
        tables vName;
        title "Frequency of Date/Time Variables Found";
    run;
%end;
%else %put WARNING: _dlib.datelist was not created;

/* Step 7: Cleanup */
libname _dlib clear;
%sysexec rmdir /s /q "&tmppath";
proc datasets lib=work nolist;
    delete _ae _dm;
quit;
