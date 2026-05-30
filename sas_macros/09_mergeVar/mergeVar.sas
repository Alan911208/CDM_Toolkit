/*****************************************************************************
 * Tool 9:  %mergeVar
 * Purpose: Merge variables from one "source" form dataset into every other
 *          dataset in the library, using specified key variables.
 *
 * Parameters:
 *   form = Source form dataset name
 *   vars = Variable(s) to merge (space-separated)
 *   key  = Key variable(s) for the merge join (e.g., usubjid)
 *   src  = Source library path
 *   tgt  = Output library path
 *
 * Usage:
 *   %mergeVar(form=dm1, vars=rficdtc, key=usubjid, src=D:\raw, tgt=D:\enriched);
 *   %mergeVar(form=dm, vars=rficdtc rfendtc, key=usubjid, src=C:\data, tgt=C:\out);
 *
 * Source: DM_Toolbox.sas lines 631-659
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

%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;

%macro mergeVar(form=, vars=, key=, src=, tgt=);
    %local _dslist _k _dsname;
    libname _r1 "&src";
    libname _r2 "&tgt";

    /* Extract the variables to merge */
    data _a;
        set _r1.&form;
        keep &key &vars;
    run;
    proc sort data=_a; by &key; run;

    %_getDatasetList(libref=_r1, outvar=_dslist);
    %let _k = 1;
    %do %while (%scan(&_dslist, &_k, %str( )) ne );
        %let _dsname = %scan(&_dslist, &_k, %str( ));
        proc sort data=_r1.&_dsname out=__src; by &key; run;
        data _r2.&_dsname;
            merge __src _a;
            by &key;
        run;
        %let _k = %eval(&_k + 1);
    %end;

    %_dropIfExists(dsn=work._a);
    %_dropIfExists(dsn=work.__src);
    libname _r1 clear;
    libname _r2 clear;
%mend mergeVar;
