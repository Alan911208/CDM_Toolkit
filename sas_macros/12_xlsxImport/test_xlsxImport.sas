/***************************************************************************/
/*  Test: xlsxImport                                                        */
/*  Description: Well-commented template test for xlsxImport.               */
/*               Since a real .xlsx file cannot be created from SAS alone,  */
/*               this script assumes a pre-existing workbook exists.        */
/*                                                                          */
/*  Prerequisites:                                                          */
/*  1. SAS/ACCESS to PC Files license must be available                    */
/*  2. A test workbook must exist at the path below, containing:            */
/*     - Sheet "Demog" with columns: SUBJID, AGE, SEX, RACE                */
/*     - Sheet "AE"    with columns: SUBJID, AETERM, AESTDT, AESEV         */
/*     Each sheet should have headers in row 1 and at least 3 data rows.   */
/***************************************************************************/

/* Step 1: Include the macro */
%include "E:\Project\Tools_DM\sas_macros\12_xlsxImport\xlsxImport.sas";

/* Step 2: Set up target library */
/* NOTE: Change path to an existing writable directory */
%let _outdir = %sysfunc(pathname(work));
libname _xlib "&_outdir";

/* Step 3: Define path to the test Excel workbook */
/* NOTE: Update this path to point to your test .xlsx file */
%let _xlfile = C:\temp\test_2sheets.xlsx;

/* Step 4: Import all sheets */
%xlsxImport(file=&_xlfile, out=_xlib);

/* Step 5: Verify results */
proc datasets lib=_xlib nolist nowarn;
quit;

title "xlsxImport — Sheet: Demog";
proc print data=_xlib.Demog noobs label;
run;

title "xlsxImport — Sheet: AE";
proc print data=_xlib.AE noobs label;
run;

title "xlsxImport — Sheet: Demog — Contents";
proc contents data=_xlib.Demog;
run;

title "xlsxImport — Sheet: AE — Contents";
proc contents data=_xlib.AE;
run;
title;

/* Step 6: Cleanup */
proc datasets lib=_xlib nolist nowarn kill;
quit;
libname _xlib clear;
