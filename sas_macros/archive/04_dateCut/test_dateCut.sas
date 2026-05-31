/*****************************************************************************
 *  Test Script: %dateCut
 *  Creates source library with AE (AESTDTC + visit var) and DM (demographics)
 *  datasets, then runs dateCut with a cutoff date.
 *  Note: This test uses simple date variables to keep the example clear.
 *  The real macro uses numeric SAS date variables (days since 1960-01-01),
 *  so we create numeric YMD date values.
 *****************************************************************************/

/* Include the macro definition (includes internal helper _dropIfExists) */
%include "E:\Project\Tools_DM\sas_macros\04_dateCut\dateCut.sas";

/* Step 1: Create temporary directories for source and target */
%let srcpath = %sysfunc(pathname(work))\datesrc;
%let tgtpath = %sysfunc(pathname(work))\datetgt;
%sysexec mkdir "&srcpath";
%sysexec mkdir "&tgtpath";

/* Step 2: Create AE (Adverse Events) dataset with date variables
   USUBJID, Visit, AESTDTC_YMD (numeric SAS date), AETERM
   AESTDTC_YMD = days since 1960-01-01, so 2024-01-15 = 23391 */
data work._ae;
    input USUBJID $ Visit AESTDTC_YMD AETERM $ &;
    datalines;
001 1   23391 Headache
001 2   23400 Nausea
001 3   23415 Fatigue
002 1   23385 Dizziness
002 2   23410 Vertigo
003 1   23420 Pain
;
run;
/* Verify: 23391 = '15JAN2024'd, 23400 = '24JAN2024'd, etc. */

/* Step 3: Create DM (Demographics) dataset
   USUBJID, SEX, AGE, ARM, RFICDTC_YMD */
data work._dm;
    input USUBJID $ SEX $ AGE ARM $ RFICDTC_YMD;
    datalines;
001 M 45 TRT 23380
002 F 52 CTL 23382
003 F 38 TRT 23405
;
run;

/* Step 4: Save datasets to the source library */
libname _src "&srcpath";
libname _tgt "&tgtpath";
data _src.ae; set work._ae; run;
data _src.dm; set work._dm; run;

/* Step 5: Check pre-cut record counts */
%put === Source library record counts ===;
%let _n_ae = 6;   /* expected 6 rows in AE */
%let _n_dm = 3;   /* expected 3 rows in DM */

/* Step 6: Run dateCut
   Cutoff = 2024-01-20 (numeric value = 23396, i.e. '20JAN2024'd)
   Visit variable = AE.AESTDTC_YMD (visit dates from AE dataset)
   otherVar = AE.AESTDTC_YMD (direct date filter on AE.AESTDTC_YMD)
   Note: We use a simple case where visitVar matches the dataset being filtered */
%put === Running dateCut with cutoff=20240120 ===;
%dateCut(
    src      = &srcpath,
    tgt      = &tgtpath,
    cutoff   = 20240120,
    visitVar = AE.AESTDTC_YMD,
    otherVar = DM.RFICDTC_YMD,
    keepAll  =,
    byVar    = USUBJID Visit
);

/* Step 7: Print the comparison report */
%if %sysfunc(exist(_tgt._rpt_, data)) %then %do;
    proc print data=_tgt._rpt_ label noobs;
        title "dateCut Comparison Report";
    run;
%end;
%else %put WARNING: _tgt._rpt_ was not created;

/* Step 8: Print filtered datasets */
%if %sysfunc(exist(_tgt.ae, data)) %then %do;
    proc print data=_tgt.ae noobs;
        title "Filtered AE Dataset (visits before 2024-01-20)";
    run;
%end;
%if %sysfunc(exist(_tgt.dm, data)) %then %do;
    proc print data=_tgt.dm noobs;
        title "Filtered DM Dataset (RFICDTC_YMD before 2024-01-20)";
    run;
%end;

/* Step 9: Cleanup */
libname _src clear;
libname _tgt clear;
%sysexec rmdir /s /q "&srcpath";
%sysexec rmdir /s /q "&tgtpath";
proc datasets lib=work nolist;
    delete _ae _dm;
quit;
