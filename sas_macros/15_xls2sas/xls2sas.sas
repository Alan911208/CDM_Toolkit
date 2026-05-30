/***************************************************************************/
/*  Macro: xls2sas                                                          */
/*  Purpose: Import a specified sheet from an Excel workbook into a SAS     */
/*           dataset. A convenience wrapper around PROC IMPORT with         */
/*           sensible defaults and parameter validation.                    */
/*  Parameters:                                                              */
/*    file    = Full path to the Excel workbook (required)                  */
/*    insheet = Sheet name to import (required)                             */
/*    outset  = Output SAS dataset name (required)                          */
/*    outlib  = Output library (default: work)                              */
/*    dbms    = File type: xlsx, xls, or excel (default: xlsx)             */
/*    getnames= Treat first row as variable names? YES or NO (default: YES)*/
/*    mixed   = Allow mixed char/num in a column? YES or NO (default: NO)  */
/*    usedate = Use SAS date format for date columns? YES or NO (def: YES) */
/*    scantext= Scan text columns for max length? YES or NO (default: YES) */
/*                                                                          */
/*  Dependencies: SAS/ACCESS to PC Files (for Excel import)                 */
/*  Author: DM Tools                                                        */
/*  Version: 1.0                                                            */
/***************************************************************************/

%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;

%macro xls2sas(file=, insheet=, outset=, outlib=work, dbms=xlsx,
               getnames=YES, mixed=NO, usedate=YES, scantext=YES);

    /* Validate required parameters */
    %if %length(&file) = 0 %then %do;
        %put ERROR: xls2sas - file parameter is required;
        %return;
    %end;

    %if %length(&insheet) = 0 %then %do;
        %put ERROR: xls2sas - insheet parameter is required;
        %return;
    %end;

    %if %length(&outset) = 0 %then %do;
        %put ERROR: xls2sas - outset parameter is required;
        %return;
    %end;

    /* Check if the Excel file exists on disk */
    %local _fid _rc _exist;
    %let _exist = 0;
    data _null_;
        length _f $512;
        _f = "&file";
        _fid = fopen(_f, 'i', 0);
        if _fid > 0 then do;
            call symputx("_exist", 1);
            _rc = fclose(_fid);
        end;
    run;

    %if &_exist = 0 %then %do;
        %put ERROR: xls2sas - File not found: &file;
        %return;
    %end;

    /* Drop existing output dataset if it exists */
    %_dropIfExists(dsn=&outlib..&outset);

    %put NOTE: xls2sas - Importing sheet "&insheet" from &file;
    %put NOTE: xls2sas - Output dataset: &outlib..&outset;

    proc import datafile="&file"
        out=&outlib..&outset
        dbms=&dbms replace;
        sheet="&insheet";
        getnames=&getnames;
        mixed=&mixed;
        usedate=&usedate;
        scantext=&scantext;
    run;

    %put NOTE: xls2sas - Import complete. Dataset &outlib..&outset created.;
%mend xls2sas;
