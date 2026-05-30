/*****************************************************************************
 * Internal Helper: _getDatasetList
 * Purpose:  Populate a macro variable with dataset names from a library.
 * Parameters:
 *   libref  = SAS libref to scan
 *   outvar  = Name of macro variable to hold the space-delimited list
 *   exclude = Optional dataset name to exclude
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


/*****************************************************************************
 * Tool 17: %runDateDiff
 * Purpose:  Compare old-vs-new of a single dataset at the observation level.
 *           Only changed rows get today's date; unchanged rows keep the
 *           original rundate.
 *
 * Parameters:
 *   oldlib / oldset = Old (baseline) library and dataset name
 *   newlib / newset = New (current) library and dataset name
 *   outlib / outset = Output library and dataset name
 *   keyvar           = Key variables to match rows (space-separated)
 *   exvar            = Variables to exclude from the comparison
 *
 * Notes:
 *   - Both oldset and newset must exist in their respective libraries.
 *   - The new dataset must contain variable 'rundate'.
 *   - exvar should include 'rundate' to prevent comparing the date itself.
 *****************************************************************************/
%macro runDateDiff(oldlib=, oldset=, newlib=, newset=, outlib=, outset=, keyvar=, exvar=);
    %local _vars _cvarnum _ii;

    proc sql noprint;
        select name into :_vars separated by ' '
        from dictionary.columns
        where upcase(libname) = "%upcase(&newlib)"
          and upcase(memname) = "%upcase(&newset)" and memtype = "DATA"
          and upcase(name) not in (%upcase("&keyvar &exvar"));
    quit;

    %let _cvarnum = %sysfunc(countw(&_vars, %str( )));

    proc sort data=&newlib..&newset; by &keyvar; run;
    proc sort data=&oldlib..&oldset; by &keyvar; run;

    data &outlib..&outset;
        merge &oldlib..&oldset(
                rename=(
                    %do _ii = 1 %to &_cvarnum;
                        %scan(&_vars, &_ii, %str( )) = _%scan(&_vars, &_ii, %str( ))
                    %end;
                    rundate = _rundate)
                in=_zOld)
              &newlib..&newset(in=_zNew);
        by &keyvar;
        if _zNew;
        if _zOld
            %do _ii = 1 %to &_cvarnum;
                and %scan(&_vars, &_ii, %str( )) = _%scan(&_vars, &_ii, %str( ))
            %end;
        then rundate = _rundate;
        else rundate = "&sysdate9.";
        drop %do _ii = 1 %to &_cvarnum;
                _%scan(&_vars, &_ii, %str( ))
             %end;
             _rundate;
    run;
%mend runDateDiff;


/*****************************************************************************
 * Tool 18: %batchRunDateDiff
 * Purpose:  Apply %runDateDiff to every dataset in a library.
 *
 * Parameters:
 *   newLib = Path to new (current) data
 *   oldLib = Path to old (baseline) data
 *   outLib = Path to output data
 *   keyvar = Key variables for row matching (space-separated)
 *   exvar  = Variables to exclude from comparison
 *
 * Usage:
 *   %batchRunDateDiff(newLib=D:\new, oldLib=D:\old, outLib=D:\final,
 *       keyvar=usubjid PAGE LINE, exvar=STUDYID SITE FORM rundate);
 *
 * Logic:
 *   1. Assign librefs _new, _old, _out to the three paths.
 *   2. Get the list of datasets present in _old via _getDatasetList.
 *   3. For each dataset, call %runDateDiff with the same dataset name
 *      for oldset, newset, and outset.
 *   4. Clear all temporary librefs.
 *
 * Notes:
 *   - Each dataset in oldLib must also exist in newLib with the same name.
 *   - Each dataset should already have a 'rundate' field (e.g. set via
 *     %stampRunDate).
 *   - The output library must be writeable.
 *****************************************************************************/
%macro batchRunDateDiff(newLib=, oldLib=, outLib=, keyvar=, exvar=);
    %local _dslist _k _dsname;
    libname _new  "&newLib";
    libname _old  "&oldLib";
    libname _out  "&outLib";

    %_getDatasetList(libref=_old, outvar=_dslist);

    %let _k = 1;
    %do %while (%scan(&_dslist, &_k, %str( )) ne );
        %let _dsname = %scan(&_dslist, &_k, %str( ));
        %runDateDiff(oldlib=_old, oldset=&_dsname,
                     newlib=_new, newset=&_dsname,
                     outlib=_out, outset=&_dsname,
                     keyvar=&keyvar, exvar=&exvar);
        %let _k = %eval(&_k + 1);
    %end;
    libname _new clear; libname _old clear; libname _out clear;
%mend batchRunDateDiff;
