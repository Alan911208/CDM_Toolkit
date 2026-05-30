/***************************************************************************/
/*  Test: xls2sas                                                           */
/*  Description: Template test for xls2sas. Since a real .xlsx/.xls file    */
/*               cannot be created from SAS alone, this script demonstrates */
/*               how to call xls2sas with a pre-existing Excel file.        */
/*                                                                          */
/*  Prerequisites:                                                          */
/*  1. SAS/ACCESS to PC Files license                                      */
/*  2. A test Excel workbook at the path below with:                        */
/*     - Sheet "lb10" with columns: SUBJID, LBTEST, LBRES, LBUNIT          */
/*     - Row 1: variable names (headers)                                    */
/*     - Rows 2-5: data                                                     */
/***************************************************************************/

/* Step 1: Include the macro */
%include "E:\Project\Tools_DM\sas_macros\15_xls2sas\xls2sas.sas";

/* Step 2: Define the path to the test Excel file */
/* NOTE: Update to point to your test .xlsx or .xls file */
%let _xlsfile = C:\temp\test_lab_data.xlsx;

/* Step 3: Import with basic parameters */
%xls2sas(file=&_xlsfile, insheet=lb10, outset=lb10_imported);

/* Step 4: Verify the imported dataset */
title "xls2sas — Imported dataset: lb10_imported";
proc print data=work.lb10_imported noobs label;
run;

proc contents data=work.lb10_imported;
run;
title;

/* Step 5: Test with additional options */

/* Import without treating row 1 as headers (variable names will be F1, F2, ...) */
%xls2sas(file=&_xlsfile, insheet=lb10, outset=lb10_nohead, getnames=NO);

title "xls2sas — Imported without headers";
proc print data=work.lb10_nohead noobs;
run;
title;

/* Step 6: Check that the original imported dataset has:
   - Correct number of observations matching the Excel data rows
   - Variable names matching the Excel header row
   - Appropriate variable types (character/numeric) */
proc datasets lib=work nolist nowarn;
quit;

/* Step 7: Cleanup */
proc datasets lib=work nolist nowarn;
    delete lb10_imported lb10_nohead;
quit;
