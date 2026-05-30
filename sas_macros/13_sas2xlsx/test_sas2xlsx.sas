/***************************************************************************/
/*  Test: sas2xlsx                                                          */
/*  Description: Creates 3 small datasets (demog, ae, lb) in WORK, then     */
/*               exports them all to a single .xlsx workbook.               */
/*                                                                          */
/*  Prerequisites:                                                          */
/*  1. SAS/ACCESS to PC Files license (for PROC EXPORT dbms=xlsx)          */
/*  2. Write access to the output directory                                */
/***************************************************************************/

/* Step 1: Include the macro */
%include "E:\Project\Tools_DM\sas_macros\13_sas2xlsx\sas2xlsx.sas";

/* Step 2: Create mock datasets in WORK */

/* Dataset 1: demog */
data work.demog;
    label subjid = "Subject ID"  age = "Age (years)"  sex = "Sex";
    input subjid $ age sex $;
    datalines;
101 45 M
102 38 F
103 52 F
104 41 M
    ;
run;

/* Dataset 2: ae */
data work.ae;
    label subjid = "Subject ID"  aeterm = "AE Term"  aestdt = "AE Start Date";
    input subjid $ aeterm $20. aestdt yymmdd10.;
    format aestdt yymmdd10.;
    datalines;
101 Headache 2023-01-15
102 Nausea   2023-02-20
103 Rash     2023-03-05
    ;
run;

/* Dataset 3: lb */
data work.lb;
    label subjid = "Subject ID"  lbtest = "Lab Test"  lbres = "Result";
    input subjid $ lbtest $ lbres;
    datalines;
101 ALT 25
101 AST 30
102 ALT 18
102 AST 22
    ;
run;

/* Step 3: Export all datasets to a single .xlsx workbook */
%let _outdir = %sysfunc(pathname(work));

%sas2xlsx(lib=work, dir=&_outdir, name=test_export, ext=xlsx);

/* Step 4: Verify output file exists */
data _null_;
    length _fname $256;
    _fname = "&_outdir\test_export.xlsx";
    _fid = fopen(_fname, 'i', 0);
    if _fid > 0 then do;
        put "NOTE: Output file created successfully: " _fname;
        _rc = fclose(_fid);
    end;
    else put "WARNING: Output file not found: " _fname;
run;

/* Step 5: Cleanup */
proc datasets lib=work nolist nowarn;
    delete demog ae lb;
quit;
