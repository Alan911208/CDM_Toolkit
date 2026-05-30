/*****************************************************************************
 * Macro:      _getDatasetList (internal helper)
 * Purpose:    Populate a macro variable with dataset names from a library.
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
 * Tool 16: %stampRunDate
 * Purpose:  Assign a fixed run-date value to every dataset in a library.
 *
 * Parameters:
 *   lib  = Path to the SAS data library (physical path or libref)
 *   date = SAS date literal, e.g. '07JUL2020'd or "&sysdate9."d
 *
 * Usage:
 *   %stampRunDate(lib=C:\delivery, date='07JUL2020'd);
 *
 * Notes:
 *   - A temporary libref (_tmp1) is assigned to the path, then cleared.
 *   - Each dataset in the library receives a new variable named 'rundate'.
 *   - If 'rundate' already exists, it is overwritten with the supplied value.
 *   - The date parameter must be a valid SAS date literal (number of days
 *     since 01JAN1960) or a date constant with the 'd' suffix.
 *****************************************************************************/
%macro stampRunDate(lib=, date=);
    %local _dslist _k _dsname;
    libname _tmp1 "&lib";
    %_getDatasetList(libref=_tmp1, outvar=_dslist);

    %let _k = 1;
    %do %while (%scan(&_dslist, &_k, %str( )) ne );
        %let _dsname = %scan(&_dslist, &_k, %str( ));
        data _tmp1.&_dsname;
            set _tmp1.&_dsname;
            rundate = &date;
        run;
        %let _k = %eval(&_k + 1);
    %end;
    libname _tmp1 clear;
%mend stampRunDate;
