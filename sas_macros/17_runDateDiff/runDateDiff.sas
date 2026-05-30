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
 * Usage:
 *   %runDateDiff(oldlib=old, oldset=_mh2, newlib=new, newset=_mh2,
 *       outlib=out, outset=rst, keyvar=usubjid PAGE LINE,
 *       exvar=STUDYID SITE FORM rundate);
 *
 * Logic:
 *   1. Identify all columns in the new dataset, excluding keyvar and exvar.
 *   2. Merge old and new datasets by keyvar.  Old columns are renamed with
 *      a leading underscore.  The old rundate is renamed to _rundate.
 *   3. For each observation in new:
 *        - If a matching row exists in old AND all comparable columns are
 *          identical, keep the original rundate (_rundate).
 *        - Otherwise (new row or changed row), assign today's date
 *          from &sysdate9.
 *   4. Drop all the temporary underscore-prefixed columns and _rundate.
 *
 * Notes:
 *   - Both oldset and newset must exist in their respective libraries.
 *   - The new dataset must contain variable 'rundate' (likely set by
 *     stampRunDate earlier).
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
