/*****************************************************************************
 * Test Script: reviewData
 * Purpose: Template showing how to set up a data library and run the
 *          interactive data review.  This is a TEMPLATE — it requires
 *          Windows and Excel with DDE enabled to execute fully.
 *
 * WARNING: This macro REQUIRES:
 *   1. Windows operating system
 *   2. Microsoft Excel installed
 *   3. DDE enabled in Excel:
 *      File > Options > Advanced > General >
 *      "Ignore other applications that use Dynamic Data Exchange (DDE)" = OFF
 *   4. Recommended: 32-bit SAS + 32-bit Excel
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
    %include "&_projectRoot\sas_macros\20_reviewData\reviewData.sas";

    * ---- Step 2: Create a temporary data library ----;
    %let revDir = %sysfunc(pathname(WORK))\rev_data;
    %sysexec mkdir "&revDir";
    libname revLib "&revDir";

    * ---- Step 3: Create mock clinical datasets ----;

    * DM — demographics;
    data revLib.dm;
        length studyid $10 usubjid $10 siteid $4 armcd $8;
        studyid = "STUDY01"; usubjid = "001-0001"; siteid = "S001"; armcd = "TRT"; output;
        studyid = "STUDY01"; usubjid = "001-0002"; siteid = "S001"; armcd = "PBO"; output;
        studyid = "STUDY01"; usubjid = "001-0003"; siteid = "S002"; armcd = "TRT"; output;
    run;

    * AE — adverse events;
    data revLib.ae;
        length studyid $10 usubjid $10 aeterm $50 aesev $8;
        studyid = "STUDY01"; usubjid = "001-0001"; aeterm = "Headache";  aesev = "MILD";    output;
        studyid = "STUDY01"; usubjid = "001-0001"; aeterm = "Nausea";    aesev = "MODERATE"; output;
        studyid = "STUDY01"; usubjid = "001-0002"; aeterm = "Fatigue";   aesev = "MILD";    output;
    run;

    * MH — medical history;
    data revLib.mh;
        length studyid $10 usubjid $10 mhterm $50;
        studyid = "STUDY01"; usubjid = "001-0001"; mhterm = "Hypertension"; output;
        studyid = "STUDY01"; usubjid = "001-0002"; mhterm = "Diabetes";      output;
        studyid = "STUDY01"; usubjid = "001-0003"; mhterm = "Asthma";         output;
    run;

    * QUERY — optional query findings for color coding;
    data revLib.query;
        length tblname $10 var $32 qrytype $12 username $10;
        * Confirmed finding on AE.Headache;
        tblname = "AE"; var = "AETERM"; qrytype = "OTHER"; username = "QC";
        output;
        * Pending entry query on MH.MHTERM;
        tblname = "MH"; var = "MHTERM"; qrytype = "ENTRYqry"; username = "Pending";
        output;
        * Closed edit query on DM.ARMCD;
        tblname = "DM"; var = "ARMCD"; qrytype = "EDITqry"; username = "Closed";
        output;
    run;

    * CROSSCHK — optional crosscheck findings;
    data revLib.crosschk;
        length usubjid $10 pageid $6 dataset $20 var $32 line 8 finding $200;
        usubjid = "001-0001"; pageid = "AE01"; dataset = "AE";
        var = "AESEV"; line = 1; finding = "Data Confirmed"; output;
        usubjid = "001-0002"; pageid = "DM01"; dataset = "DM";
        var = "ARMCD"; line = 1; finding = "Supported by treatment/drugs record"; output;
    run;

    libname revLib clear;

    * ---- Step 4: Run reviewData (FULL style) ----;
    %reviewData(path=&revDir, style=FULL, showFail=NO);

    * The macro opens Excel and populates it. Leave Excel open for review.;

    * ---- Step 5: Clean up work datasets ----;
    proc datasets lib=work nolist kill;
    quit;

%mend _runTest;

%macro _testReviewData();
    %if &_DDE_SKIP = 1 %then %do;
        %put ============================================;
        %put NOTE: reviewData test SKIPPED — DDE not available.;
        %put NOTE: This macro REQUIRES Windows + Excel with DDE.;
        %put NOTE: See test_results.txt for expected behavior.;
        %put ============================================;
    %end;
    %else %do;
        %_runTest
    %end;
%mend _testReviewData;

%_testReviewData
