/*****************************************************************************
 * Test Script: compareLib
 * Purpose: Template showing how to set up old/new libraries and run the
 *          library comparison.  This is a TEMPLATE — it requires Windows
 *          and Excel with DDE enabled to execute fully.
 *
 * WARNING: This macro REQUIRES:
 *   1. Windows operating system
 *   2. Microsoft Excel installed
 *   3. DDE enabled in Excel:
 *      File > Options > Advanced > General >
 *      "Ignore other applications that use Dynamic Data Exchange (DDE)" = OFF
 *   4. Recommended: 32-bit SAS + 32-bit Excel
 *
 * This test creates two temporary libraries with small datasets, then
 * invokes compareLib.  If Excel/DDE is not available, the macro will
 * attempt to launch Excel and may hang or produce errors.
 *
 * If you cannot meet the DDE requirements, use this script as a
 * reference for the parameter setup, and review the test_results.txt
 * for the expected behavior.
 *****************************************************************************/

* ---- Step 0: Guard check — skip DDE test if requirements not met ----;
%macro _skipIfNoDDE();
    %global _DDE_SKIP;
    %let _DDE_SKIP = 0;
    filename _dde_test dde 'excel|system';
    data _null_;
        length fid 8;
        fid = fopen('_dde_test','s');
        if fid le 0 then do;
            call symput('_DDE_SKIP','1');
            put "NOTE: DDE channel not available. Skipping interactive Excel test.";
            put "NOTE: See test_results.txt for expected behavior description.";
        end;
        else rc = fclose(fid);
    run;
    filename _dde_test clear;
%mend _skipIfNoDDE;
%_skipIfNoDDE

%macro _runTest();
    * ---- Step 1: Include the macro ----;
    %include "&_projectRoot\sas_macros\19_compareLib\compareLib.sas";

    * ---- Step 2: Create temporary directories ----;
    %let cmpOld = %sysfunc(pathname(WORK))\cmp_old;
    %let cmpNew = %sysfunc(pathname(WORK))\cmp_new;
    %sysexec mkdir "&cmpOld";
    %sysexec mkdir "&cmpNew";

    * ---- Step 3: Create mock old datasets ----;

    * DM (old) — 3 subjects;
    data dm_old;
        length studyid $10 usubjid $10 siteid $4 armcd $8;
        studyid = "STUDY01"; usubjid = "001-0001"; siteid = "S001"; armcd = "TRT"; output;
        studyid = "STUDY01"; usubjid = "001-0002"; siteid = "S001"; armcd = "PBO"; output;
        studyid = "STUDY01"; usubjid = "001-0003"; siteid = "S002"; armcd = "TRT"; output;
    run;

    * AE (old) — 3 events;
    data ae_old;
        length studyid $10 usubjid $10 aeterm $50 aesev $8;
        studyid = "STUDY01"; usubjid = "001-0001"; aeterm = "Headache";  aesev = "MILD";    output;
        studyid = "STUDY01"; usubjid = "001-0001"; aeterm = "Nausea";    aesev = "MODERATE"; output;
        studyid = "STUDY01"; usubjid = "001-0002"; aeterm = "Fatigue";   aesev = "MILD";    output;
    run;

    * ---- Step 4: Create mock new datasets (with changes) ----;

    * DM (new) — 001-0002 armcd changed PBO→TRT, 001-0004 added;
    data dm_new;
        length studyid $10 usubjid $10 siteid $4 armcd $8;
        studyid = "STUDY01"; usubjid = "001-0001"; siteid = "S001"; armcd = "TRT"; output;
        studyid = "STUDY01"; usubjid = "001-0002"; siteid = "S001"; armcd = "TRT"; output;  * CHANGED;
        studyid = "STUDY01"; usubjid = "001-0003"; siteid = "S002"; armcd = "TRT"; output;
        studyid = "STUDY01"; usubjid = "001-0004"; siteid = "S003"; armcd = "PBO"; output;  * NEW;
    run;

    * AE (new) — severity changed for Nausea, new event added;
    data ae_new;
        length studyid $10 usubjid $10 aeterm $50 aesev $8;
        studyid = "STUDY01"; usubjid = "001-0001"; aeterm = "Headache";  aesev = "MILD";   output;
        studyid = "STUDY01"; usubjid = "001-0001"; aeterm = "Nausea";    aesev = "SEVERE"; output;  * CHANGED;
        studyid = "STUDY01"; usubjid = "001-0002"; aeterm = "Fatigue";   aesev = "MILD";   output;
        studyid = "STUDY01"; usubjid = "001-0002"; aeterm = "Dizziness"; aesev = "MILD";   output;  * NEW AE;
    run;

    * ---- Step 5: Copy to library directories ----;
    libname _old "&cmpOld";
    libname _new "&cmpNew";
    proc copy in=work out=_old; select dm_old ae_old; run;
    proc copy in=work out=_new; select dm_new ae_new; run;
    libname _old clear;
    libname _new clear;

    * ---- Step 6: Run compareLib ----;
    %compareLib(
        pCompare=&cmpOld,
        pBase=&cmpNew,
        kVar=studyid usubjid aeterm,
        mode=9);

    * The macro opens Excel and populates it. Leave Excel open for review.;

    * ---- Step 7: Clean up WORK datasets ----;
    proc datasets lib=work nolist;
        delete dm_old ae_old dm_new ae_new;
    quit;

    * ---- Step 8: Clean up directories (keep Excel open) ----;
    %sysexec rmdir /s /q "&cmpOld";
    %sysexec rmdir /s /q "&cmpNew";

%mend _runTest;

%macro _testCompareLib();
    %if &_DDE_SKIP = 1 %then %do;
        %put ============================================;
        %put NOTE: compareLib test SKIPPED — DDE not available.;
        %put NOTE: This macro REQUIRES Windows + Excel with DDE.;
        %put NOTE: See test_results.txt for expected behavior.;
        %put ============================================;
    %end;
    %else %do;
        %_runTest
    %end;
%mend _testCompareLib;

%_testCompareLib
