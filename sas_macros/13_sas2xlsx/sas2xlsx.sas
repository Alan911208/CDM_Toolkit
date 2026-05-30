/***************************************************************************/
/*  Macro: sas2xlsx                                                         */
/*  Purpose: Batch-export all SAS datasets in a library to a single Excel   */
/*           workbook (.xlsx), one sheet per dataset. Uses PROC EXPORT      */
/*           with dbms=xlsx.                                                */
/*  Parameters:                                                              */
/*    lib  = SAS library name whose datasets are to be exported (required) */
/*    dir  = Output directory path (required)                               */
/*    name = Output workbook filename (without extension) — required       */
/*    ext  = File extension, typically xlsx (default) or xls               */
/*  Dependencies: SAS/ACCESS to PC Files (for dbms=xlsx)                   */
/*  Author: DM Tools                                                        */
/*  Version: 1.0                                                            */
/***************************************************************************/

%macro sas2xlsx(lib=, dir=, name=, ext=xlsx);
    %local _dslist _k _dsname _lup;

    /* Validate required parameters */
    %if %length(&lib) = 0 or %length(&dir) = 0 or %length(&name) = 0 %then %do;
        %put ERROR: sas2xlsx - lib, dir, and name are required;
        %return;
    %end;

    %let _lup = %upcase(&lib);

    %put NOTE: sas2xlsx - Enumerating datasets in library &lib;

    /* Get list of all datasets in the library */
    proc sql noprint;
        select memname into :_dslist separated by ' '
        from dictionary.tables where upcase(libname) = "&_lup";
    quit;

    %if &sqlobs = 0 %then %do;
        %put WARNING: sas2xlsx - No datasets found in library &lib;
        %return;
    %end;

    %put NOTE: sas2xlsx - Found &sqlobs dataset(s) to export;

    /* Iterate over all datasets */
    %let _k = 1;
    %do %while (%scan(&_dslist, &_k, %str( )) ne );
        %let _dsname = %scan(&_dslist, &_k, %str( ));
        %put NOTE: sas2xlsx - Exporting &lib..&_dsname -> sheet &_dsname;
        proc export data=&lib..&_dsname
            outfile="&dir\&name..&ext"
            dbms=&ext replace label;
            sheet="&_dsname";
        run;
        %let _k = %eval(&_k + 1);
    %end;

    %let _k = %eval(&_k - 1);
    %put NOTE: sas2xlsx - Successfully exported &_k dataset(s) to &dir\&name..&ext;
%mend sas2xlsx;
