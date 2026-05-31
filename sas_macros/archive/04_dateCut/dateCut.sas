/*****************************************************************************
 *  Macro:  %dateCut
 *  Purpose: Filter a library of datasets by a cutoff date. Datasets with
 *           visit-date variables are filtered (keep rows before cutoff);
 *           datasets listed in keepAll are copied in full.
 *           Produces a record-count summary (_rpt_) in the target library.
 *
 *  Parameters:
 *    src      = Path to source data library
 *    tgt      = Path to output data library
 *    cutoff   = Cutoff date in YYYYMMDD (e.g., 20170228)
 *    visitVar = Visit-date variables: "DS.var DS.var ..."
 *               (DS=dataset name, var=date variable)
 *    otherVar = Other date-filtered variables in "DS.var" format
 *    keepAll  = Datasets to copy without filtering (space-separated names)
 *    byVar    = Key variables for merging (e.g., Usubjid Visit)
 *
 *  Usage:
 *    %dateCut(src=D:\raw, tgt=D:\interim, cutoff=20170228,
 *        visitVar=SV1.SVSTDTC_YMD SV2.SVSTDTC_YMD,
 *        otherVar=ae.aestdtc_ymd, keepAll=FORMATS CO,
 *        byVar=Usubjid Visit);
 *
 *  Internal helpers included: _dropIfExists
 *
 *  Source: DM_Toolbox.sas — Tool 4, Category 2
 *****************************************************************************/

/* ===== INTERNAL HELPER ===== */
%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;

/* ===== MAIN MACRO ===== */
%macro dateCut(
    src      = D:\temp\tmp,
    tgt      = D:\temp\tmp\partdata,
    cutoff   = 20170228,
    visitVar = ,
    otherVar = ,
    keepAll  = FORMATS CO,
    byVar    = Usubjid Visit);

    %local _ex _k _token _inlist _vars _dsname _bylist;

    libname _iData "&src";
    libname _oData "&tgt";

    /* Build exclusion list for the main loop */
    %let _ex = %str( );
    %let _k = 1;
    %do %while (%scan(&otherVar, &_k, %str( )) ne );
        %let _token = %scan(&otherVar, &_k, %str( ));
        %let _ex = &_ex %upcase("%scan(&_token, 1, .)"),;
        %let _k = %eval(&_k + 1);
    %end;
    %let _k = 1;
    %do %while (%scan(&keepAll, &_k, %str( )) ne );
        %let _token = %scan(&keepAll, &_k, %str( ));
        %let _ex = &_ex %upcase("%scan(&_token, 1, .)"),;
        %let _k = %eval(&_k + 1);
    %end;

    /* Cleanup temp tables from prior runs */
    %_dropIfExists(dsn=_oData._rpt_);
    %_dropIfExists(dsn=work._rpt_);
    %_dropIfExists(dsn=work._vst_);

    /* Get dataset list excluding the ones handled separately */
    proc sql noprint;
        select memname into :_inlist separated by ' '
        from dictionary.tables
        where upcase(libname) = "%upcase(_iData)" and memtype = "DATA"
        %if %length(&_ex) > 2 %then and upcase(memname) not in (&_ex);
        ;
    quit;

    /* Stack visit-date variables into _vst_ */
    %let _k = 1;
    %do %while (%scan(&visitVar, &_k, %str( )) ne );
        %let _token = %scan(&visitVar, &_k, %str( ));
        %let _dsname = %scan(&_token, 1, .);
        %if ^%sysfunc(exist(work._vst_, data)) %then %do;
            data _vst_(keep=&byVar _vst_);
                set _iData.&_dsname;
                _vst_ = %scan(&_token, 2, .);
            run;
        %end; %else %do;
            data _vst_(keep=&byVar _vst_);
                set _vst_ _iData.&_dsname(in=_b);
                if _b then _vst_ = %scan(&_token, 2, .);
            run;
        %end;
        %let _k = %eval(&_k + 1);
    %end;

    /* Keep subjects with at least one visit before cutoff */
    proc sort data=_vst_(where=(^missing(_vst_)))
        out=_svlst_(keep=%scan(&byVar, 2, %str( ))) nodupkey;
        by %scan(&byVar, 2, %str( ));
    run;

    data _vst_;
        set _vst_;
        if _vst_ < input("&cutoff", yymmdd8.);
    run;
    proc sort data=_vst_; by &byVar; run;

    /* Build uppercase byVar list for dictionary query */
    %let _bylist = ;
    %let _k = 1;
    %do %while (%scan(&byVar, &_k, %str( )) ne );
        %let _bylist = &_bylist %upcase("%scan(&byVar, &_k, %str( ))"),;
        %let _k = %eval(&_k + 1);
    %end;

    /* Filter each dataset */
    %let _k = 1;
    %do %while (%scan(&_inlist, &_k, %str( )) ne );
        %let _dsname = %scan(&_inlist, &_k, %str( ));
        proc sql noprint;
            select count(name) into :_vars
            from dictionary.columns
            where upcase(libname) = "%upcase(_iData)"
              and upcase(memname) = "%upcase(&_dsname)"
              and upcase(name) in (&_bylist);
        quit;

        %if &_vars > 0 %then %do;
            proc sort data=_iData.&_dsname out=_tmp_; by &byVar; run;
            data _oData.&_dsname;
                merge _tmp_(in=_a) _vst_(in=_b);
                by &byVar;
                if _a and _b;
                drop _vst_;
            run;
            %_dropIfExists(dsn=work._tmp_);
        %end;
        %let _k = %eval(&_k + 1);
    %end;

    %_dropIfExists(dsn=work._vst_);
    %_dropIfExists(dsn=work._svlst_);

    /* Filter "other" datasets by date */
    %let _k = 1;
    %do %while (%scan(&otherVar, &_k, %str( )) ne );
        %let _token = %scan(&otherVar, &_k, %str( ));
        data _oData.%scan(&_token, 1, .);
            set _iData.%scan(&_token, 1, .);
            if %scan(&_token, 2, .) < input("&cutoff", yymmdd8.);
        run;
        %let _k = %eval(&_k + 1);
    %end;

    /* Copy keepAll datasets */
    %let _k = 1;
    %do %while (%scan(&keepAll, &_k, %str( )) ne );
        %let _token = %scan(&keepAll, &_k, %str( ));
        data _oData.%scan(&_token, 1, .);
            set _iData.%scan(&_token, 1, .);
        run;
        %let _k = %eval(&_k + 1);
    %end;

    /* Build comparison report */
    proc sql;
        create table _tmp_ as
            select memname, libname, nlobs
            from dictionary.tables
            where upcase(libname) in ("%upcase(_iData)", "%upcase(_oData)")
              and memtype = "DATA";
    quit;
    proc sort data=_tmp_; by memname libname; run;

    data _rpt_(drop=libname1 label=' ');
        merge _tmp_(where=(libname="%upcase(_iData)")
                    rename=(libname=libname1 nlobs=nlobs0))
              _tmp_(where=(libname="%upcase(_oData)")
                    rename=(nlobs=nlobs1));
        by memname;
        if nlobs0 > 0 then pct = put(nlobs1 * 100 / nlobs0, 8.1) || "%";
        else pct = "     --";
        label memname='Dataset' nlobs0='Original N' nlobs1='New N' pct='Cut %';
    run;

    proc sort data=_rpt_ out=_oData._rpt_; by memname nlobs1; run;
    proc sql; drop table _tmp_; quit;
    %put NOTE: dateCut — summary written to &tgt.\_rpt_;
%mend dateCut;
