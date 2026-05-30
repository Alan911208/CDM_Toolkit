/*****************************************************************************
 * Tool 8:  %dropVarSuffix
 * Purpose: Batch-drop variables whose names end with a given suffix from
 *          every dataset in a library.
 *
 * Parameters:
 *   suffix = Variable name suffix to match (default: _U)
 *   src    = Source library path
 *   tgt    = Output library path
 *
 * Usage:
 *   %dropVarSuffix(suffix=_U, src=D:\raw, tgt=D:\clean);
 *   %dropVarSuffix(suffix=_OLD, src=C:\data\derived, tgt=C:\data\final);
 *
 * Source: DM_Toolbox.sas lines 584-612
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

%macro dropVarSuffix(suffix=_u, src=, tgt=);
    %local _keeplist _dslist _k _dsname;
    libname _r1 "&src";
    libname _r2 "&tgt";

    proc sql noprint;
        select distinct name into :_keeplist separated by ' '
        from dictionary.columns
        where upcase(libname) = "_R1" and upcase(name) like "%upcase(&suffix)";
    quit;

    %if %length(&_keeplist) = 0 %then %do;
        %put WARNING: dropVarSuffix — no variables with suffix &suffix found;
        %return;
    %end;

    %_getDatasetList(libref=_r1, outvar=_dslist);
    %let _k = 1;
    %do %while (%scan(&_dslist, &_k, %str( )) ne );
        %let _dsname = %scan(&_dslist, &_k, %str( ));
        data _r2.&_dsname;
            set _r1.&_dsname;
            drop &_keeplist;
        run;
        %let _k = %eval(&_k + 1);
    %end;
    libname _r1 clear;
    libname _r2 clear;
%mend dropVarSuffix;
