/***************************************************************************/
/*  Test: sas2xls                                                            */
/*  Description: Template test for sas2xls (DDE-based export).              */
/*               This test script demonstrates how to use the macro IF       */
/*               DDE is available in your SAS/Windows environment.           */
/*                                                                           */
/*  CRITICAL PREREQUISITES:                                                  */
/*  1. Must run on Windows with Microsoft Excel installed                   */
/*  2. Excel must be RUNNING before executing this script                   */
/*     (Start Excel manually, or uncomment the X command below)             */
/*  3. SAS session must have DDE capability                                */
/*     (not available in SAS/Connect, SAS Server, or UNIX SAS)             */
/*                                                                           */
/*  NOTE: If DDE is not available, use sas2xlsx (macro 13) instead.         */
/***************************************************************************/

/* Step 1: Include the macro */
%include "E:\Project\Tools_DM\sas_macros\14_sas2xls\sas2xls.sas";

/* Step 2: Create a mock dataset for export */
data work.test_ae;
    input subjid $ aeterm $20. aestdt yymmdd10.;
    format aestdt yymmdd10.;
    label subjid = "Subject ID" aeterm = "Adverse Event" aestdt = "Start Date";
    datalines;
101 Headache            2023-01-15
102 Nausea              2023-02-20
103 Dizziness           2023-03-05
104 Fatigue             2023-04-12
    ;
run;

/* Step 3 (OPTIONAL): Launch Excel via X command if not already running
   Uncomment the following line if Excel is not running:
*/
/* options noxwait noxsync; */
/* x '"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"'; */

/* Step 4: Call the macro */
%let _xlout = %sysfunc(pathname(work))\test_ae_output.xls;

/*
 * Uncomment the following sections to actually execute the DDE export.
 * The macro itself contains internal comments — the call below will
 * print the template guidance without doing the actual DDE write.
 */

/* %sas2xls(inset=test_ae, outsheet=AE_Data, outfile=&_xlout); */

/*
 * --- Full DDE export example (standalone, without the macro wrapper) ---
 * Uncomment and adjust column names/types to match your dataset:
 */

/*
* Launch Excel (if needed);
options noxwait noxsync;
x '"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"';

* Save blank workbook;
filename ddecmd dde 'Excel|System';
data _null_;
    file ddecmd;
    put '[NEW(1)]';
    put '[SAVE.AS("' "&_xlout" '")]';
run;

* Write header row;
filename ddedat dde "Excel|AE_Data!R1C1:R1C3";
data _null_;
    file ddedat;
    put 'Subject ID' '09'x 'Adverse Event' '09'x 'Start Date';
run;

* Write data rows;
filename ddedat dde "Excel|AE_Data!R2C1:R100C3";
data _null_;
    set work.test_ae;
    file ddedat;
    put subjid '09'x aeterm '09'x aestdt;
run;

* Save and close;
data _null_;
    file ddecmd;
    put '[SAVE()]';
    put '[CLOSE()]';
run;

* Verify output;
data _null_;
    length _fname $256;
    _fname = "&_xlout";
    _fid = fopen(_fname, 'i', 0);
    if _fid > 0 then do;
        put "NOTE: DDE export complete: " _fname;
        _rc = fclose(_fid);
    end;
    else put "WARNING: Output file not found: " _fname;
run;
*/

/* Step 5: Cleanup */
proc datasets lib=work nolist nowarn;
    delete test_ae;
quit;
