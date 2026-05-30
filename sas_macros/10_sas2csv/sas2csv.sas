/*****************************************************************************
 * Tool 10: %sas2csv
 * Purpose: Batch-convert every SAS dataset in a directory to an individual
 *          CSV file.
 *
 * Parameters:
 *   src = Path to SAS data directory
 *   tgt = Output directory for CSV files (default: same as src)
 *
 * Usage:
 *   %sas2csv(src=D:\Project\Derived, tgt=D:\Project\csv);
 *   %sas2csv(src=C:\data\raw);
 *
 * Source: DM_Toolbox.sas lines 678-695
 *****************************************************************************/

%macro _getDatasetList(libref=, outvar=_dslist, exclude=);
    %local _ex;
    %let _ex = %str(%upcase("FORMATS"));
    %if %length(&exclude) > 0 %then
        %let _ex = &_ex, %upcase("&exclude");
    proc sql noprint;
        select memname into :&outvar separated by ' '
        from dictionary.tables
        where upcase(libname) = "%upcase(&libref)" and memtype = "DATA"
          and upcase(memname) not in (&_ex);
    quit;
%mend _getDatasetList;

%macro sas2csv(src=, tgt=);
    %local _dslist _k _dsname;
    %if %length(&tgt) = 0 %then %let tgt = &src;
    libname _inlib "&src";
    options nofmterr fmtsearch=(_inlib.formats) notes;

    %_getDatasetList(libref=_inlib, outvar=_dslist);
    %let _k = 1;
    %do %while (%scan(&_dslist, &_k, %str( )) ne );
        %let _dsname = %scan(&_dslist, &_k, %str( ));
        proc export data=_inlib.&_dsname
            outfile="&tgt.\&_dsname..csv"
            dbms=csv replace label;
        run;
        %let _k = %eval(&_k + 1);
    %end;
    libname _inlib clear;
%mend sas2csv;
