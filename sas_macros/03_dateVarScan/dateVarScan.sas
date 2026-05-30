/*****************************************************************************
 *  Macro:  %dateVarScan
 *  Purpose: Scan all datasets in a SAS library for date/time-related
 *           character variables (names containing DTC, DAT, or TIM).
 *           Results are written to libref.datelist.
 *
 *  Parameters:
 *    libref = Library reference name pointing to the data directory (required)
 *
 *  Usage:
 *    %dateVarScan(libref=mylib);
 *    proc print data=mylib.datelist; run;
 *
 *  Internal helpers included: _dropIfExists
 *
 *  Source: DM_Toolbox.sas — Tool 3, Category 1
 *****************************************************************************/

/* ===== INTERNAL HELPER ===== */
%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;

/* ===== MAIN MACRO ===== */
%macro dateVarScan(libref=);
    %local _inlist _k _token _dsname;
    %if %length(&libref) = 0 %then %do;
        %put ERROR: dateVarScan — libref is required;
        %return;
    %end;
    %_dropIfExists(dsn=work.datelist);

    proc sql noprint;
        select memname into :_inlist separated by ' '
        from dictionary.tables
        where upcase(libname) = "%upcase(&libref)" and memtype = "DATA"
          and upcase(memname) ne "FORMATS";
    quit;

    %if %length(&_inlist) = 0 %then %do;
        %put WARNING: dateVarScan — no datasets found in &libref;
        %return;
    %end;

    %let _k = 1;
    %do %while (%scan(&_inlist, &_k, %str( )) ne );
        %let _dsname = %scan(&_inlist, &_k, %str( ));
        data _date(keep=studyid--line vName vLabel value);
            set &libref..&_dsname;
            length vName $32 vLabel $200 value $200;
            array _c (*) _CHARACTER_;
            do _i = 1 to dim(_c);
                _vup = upcase(vname(_c(_i)));
                if (index(_vup,'DTC') or index(_vup,'DAT') or index(_vup,'TIM'))
                   and ^index(substr(_vup,1,1)||'_','_')
                then do;
                    vName  = vname(_c(_i));
                    vLabel = vlabel(_c(_i));
                    value  = _c(_i);
                    if ^missing(value) then output;
                end;
            end;
        run;

        proc sort data=_date nodupkey; by studyid--value; run;

        data datelist;
            set %if %sysfunc(exist(work.datelist,data)) %then work.datelist; _date;
        run;
        %_dropIfExists(dsn=work._date);
        %let _k = %eval(&_k + 1);
    %end;

    data &libref..datelist;
        set work.datelist;
    run;
    %_dropIfExists(dsn=work.datelist);
    %put NOTE: dateVarScan — output written to &libref..datelist;
%mend dateVarScan;
