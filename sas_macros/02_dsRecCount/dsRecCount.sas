/*****************************************************************************
 *  Macro:  %dsRecCount
 *  Purpose: Print record and variable counts for every dataset in a library.
 *           Results are written to work.jilu.
 *
 *  Parameters:
 *    lib   = Path to a directory containing SAS datasets (required)
 *    label = Optional label for the output dataset (default: blank)
 *
 *  Usage:
 *    %dsRecCount(lib=D:\data\raw, label=Raw Data Counts);
 *    proc print data=work.jilu; run;
 *
 *  Source: DM_Toolbox.sas — Tool 2, Category 1
 *****************************************************************************/

%macro dsRecCount(lib=, label=);
    %local _n;
    %if %length(&lib) = 0 %then %do;
        %put ERROR: dsRecCount — lib path is required;
        %return;
    %end;
    libname _rc "&lib";
    proc sql noprint;
        create table _rc_alldata as
            select memname, varnum, nobs
            from dictionary.tables
            where upcase(libname) = "_RC";
    quit;
    data work.jilu(drop=varnum nobs);
        length e $200;
        set _rc_alldata;
        by memname;
        if last.memname;
        e = cats(memname, "  Records:", nobs, ", Variables:", varnum);
        label e = "&label";
    run;
    proc sql noprint; drop table _rc_alldata; quit;
    libname _rc clear;
    %put NOTE: dsRecCount — output in work.jilu;
%mend dsRecCount;
