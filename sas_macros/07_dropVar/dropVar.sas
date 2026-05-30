/*****************************************************************************
 * Tool 7:  %dropVar
 * Purpose: Batch-drop the same variable(s) from every dataset in a library.
 *
 * Parameters:
 *   var = Variable(s) to drop (space-separated)
 *   src = Source library path
 *   tgt = Output library path
 *
 * Usage:
 *   %dropVar(var=__STUDYOID, src=D:\raw, tgt=D:\clean);
 *   %dropVar(var=VAR1 VAR2, src=C:\data\derived, tgt=C:\data\final);
 *
 * Source: DM_Toolbox.sas lines 551-568
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

%macro dropVar(var=, src=, tgt=);
    %local _dslist _k _dsname;
    libname _r1 "&src";
    libname _r2 "&tgt";
    %_getDatasetList(libref=_r1, outvar=_dslist);

    %let _k = 1;
    %do %while (%scan(&_dslist, &_k, %str( )) ne );
        %let _dsname = %scan(&_dslist, &_k, %str( ));
        data _r2.&_dsname;
            set _r1.&_dsname;
            drop &var;
        run;
        %let _k = %eval(&_k + 1);
    %end;
    libname _r1 clear;
    libname _r2 clear;
%mend dropVar;
