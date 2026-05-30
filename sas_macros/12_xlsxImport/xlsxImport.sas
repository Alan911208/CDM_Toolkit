/***************************************************************************/
/*  Macro: xlsxImport                                                       */
/*  Purpose: Import all sheets from an Excel workbook into individual SAS   */
/*           datasets in a target library. Uses LIBNAME engine to           */
/*           enumerate sheet names, then PROC IMPORT for each sheet.        */
/*  Parameters:                                                              */
/*    file = Full path to the Excel workbook (.xlsx) — required            */
/*    out  = Target SAS library name — required                            */
/*  Dependencies: SAS/ACCESS to PC Files (for LIBNAME xlsx engine and       */
/*                PROC IMPORT dbms=xlsx)                                    */
/*  Author: DM Tools                                                        */
/*  Version: 1.0                                                            */
/***************************************************************************/

%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;

%macro xlsxImport(file=, out=);
    %local _sheetlist _sheet1list _num _i _sheet _sheet1 _fref;

    /* Validate required parameters */
    %if %length(&file) = 0 or %length(&out) = 0 %then %do;
        %put ERROR: xlsxImport - file and out are required;
        %return;
    %end;

    /* Assign a temporary libref to the Excel workbook */
    %let _fref = _xlsim;

    %put NOTE: xlsxImport - Reading sheet list from &file;

    libname &_fref "&file";
    libname _out "&out";

    /* Query sheet names from dictionary.tables
       memname includes the trailing $ for named ranges;
       compress removes it to get the clean sheet name */
    proc sql noprint;
        create table _excel as
            select distinct compress(memname, '$') as sheet1, memname as sheet
            from dictionary.tables
            where upcase(libname) = "%upcase(&_fref)";
        select sheet, sheet1, count(*)
        into :_sheetlist separated by ' ',
             :_sheet1list separated by ' ',
             :_num
        from _excel;
    quit;

    libname &_fref clear;

    %put NOTE: xlsxImport - Found &_num sheet(s);

    /* Import each sheet */
    %do _i = 1 %to &_num;
        %let _sheet  = %scan(&_sheetlist,  &_i, %str( ));
        %let _sheet1 = %scan(&_sheet1list, &_i, %str( ));

        %put NOTE: xlsxImport - Importing sheet: &_sheet;

        proc import datafile="&file"
            out=_out.&_sheet1
            dbms=xlsx replace;
            sheet="&_sheet";
            getnames=yes;
        run;
    %end;

    /* Cleanup */
    %_dropIfExists(dsn=work._excel);
    libname _out clear;

    %put NOTE: xlsxImport - Successfully imported &_num sheet(s) to library &out;
%mend xlsxImport;
