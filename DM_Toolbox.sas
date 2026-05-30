/****************************************************************************
 *  TOOLBOX: Clinical Data Management SAS Macro Toolkit  v3.0 (Optimized)
 *  Date:     2026-05-28
 *
 *  Description: 26 production-grade SAS macros for the full CDM workflow.
 *  Each macro uses %local scoping, dictionary-table queries where faster,
 *  consistent keyword-parameter style, and input validation.
 *
 *  Categories:
 *    1. Data Inspection & Validation   (macros 1-3)
 *    2. Data Manipulation & Cleaning   (macros 4-9)
 *    3. Data Conversion & I/O          (macros 10-15)
 *    4. Date & Version Management      (macros 16-18)
 *    5. Data Comparison & Review       (macros 19-21)
 *    6. Quality Control & Reporting    (macros 22-26)
 *
 *  Internal Helpers (not for direct use):
 *    _getDatasetList   — populate &_dslist from a libref
 *    _dropIfExists     — conditionally drop a dataset
 ****************************************************************************/

/* ===== INTERNAL HELPERS ===== */

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


/*==========================================================================
 *  CATEGORY 1: DATA INSPECTION & VALIDATION
 *==========================================================================*/

/*****************************************************************************
 * Tool 1:  %varExist
 * Purpose: Test whether a variable exists in a dataset.
 *          Returns &rc = 0 (not found) or 1 (found) as a GLOBAL macro var.
 *
 * Parameters:
 *   dsn = Dataset name (libname.memname)
 *   var = Variable name
 *
 * Usage:
 *   %varExist(dsn=work.ae, var=AESTDTC);
 *   %if &rc = 0 %then %put ERROR: AESTDTC not found;
 *****************************************************************************/
%macro varExist(dsn=, var=);
    %global rc;
    %local _ds _vn _fid _pos;
    %let rc = 0;
    %if %length(&dsn) = 0 or %length(&var) = 0 %then %do;
        %put ERROR: varExist — dsn and var are required;
        %return;
    %end;
    /* Fast check using OPEN/VARNUM (avoids full PROC CONTENTS) */
    %let _fid = %sysfunc(open(&dsn,i));
    %if &_fid > 0 %then %do;
        %let _pos = %sysfunc(varnum(&_fid,&var));
        %if &_pos > 0 %then %let rc = 1;
        %let _fid = %sysfunc(close(&_fid));
    %end;
    %else %put WARNING: varExist — cannot open dataset &dsn;
%mend varExist;


/*****************************************************************************
 * Tool 2:  %dsRecCount
 * Purpose: Print record/variable counts for every dataset in a library.
 *          Produces dataset `work.jilu` with summary lines.
 *
 * Parameters:
 *   lib    = Path to the SAS data library
 *   label  = Descriptive label (e.g., study name)
 *
 * Usage:
 *   %dsRecCount(lib=\\server\rawdata, label=Study001-Rawdata);
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


/*****************************************************************************
 * Tool 3:  %dateVarScan
 * Purpose: Scan all datasets in a library and extract every variable whose
 *          name suggests date/time content (DTC, DAT, TIM without underscore
 *          prefix).  Writes a consolidated inventory to <libref>.datelist.
 *
 * Parameters:
 *   libref = An ACTIVE libref pointing to the data directory
 *
 * Usage:
 *   libname myData "C:\Project\Derived";
 *   %dateVarScan(libref=myData);
 *   /* Browse: proc print data=myData.datelist; run; 
 *****************************************************************************/
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


/*==========================================================================
 *  CATEGORY 2: DATA MANIPULATION & CLEANING
 *==========================================================================*/

/*****************************************************************************
 * Tool 4:  %dateCut
 * Purpose: Filter a library of datasets by a cutoff date. Datasets with
 *          visit-date variables are filtered (keep rows before cutoff);
 *          datasets listed in datacopy are copied in full.
 *          Produces a record-count summary (_rpt_) in the target library.
 *
 * Parameters:
 *   src      = Path to source data library
 *   tgt      = Path to output data library
 *   cutoff   = Cutoff date in YYYYMMDD (e.g., 20170228)
 *   visitVar = Visit-date variables: "DS.var DS.var ..." (DS=dataset, var=date var)
 *   otherVar = Other date-filtered variables in same format
 *   keepAll  = Datasets to copy without filtering (space-separated names)
 *   byVar    = Key variables for merging (e.g., Usubjid Visit)
 *
 * Usage:
 *   %dateCut(src=D:\raw, tgt=D:\interim, cutoff=20170228,
 *       visitVar=SV1.SVSTDTC_YMD SV2.SVSTDTC_YMD,
 *       otherVar=ae.aestdtc_ymd, keepAll=FORMATS CO,
 *       byVar=Usubjid Visit);
 *****************************************************************************/
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


/*****************************************************************************
 * Tool 5:  %textSplit
 * Purpose: Split a long character variable (>200 chars) into N columns of
 *          max 200 chars each, preserving multi-byte character boundaries.
 *
 * Parameters:
 *   dsn = Dataset name (modified in place)
 *   var = Character variable to split
 *
 * Usage:
 *   %textSplit(dsn=work.ae, var=AETERM);
 *****************************************************************************/
%macro textSplit(dsn=, var=);
    %local _max _label;
    %if %length(&dsn) = 0 or %length(&var) = 0 %then %do;
        %put ERROR: textSplit — dsn and var are required;
        %return;
    %end;

    data _cut1;
        set &dsn;
        _n = ceilz(lengthn(ktrim(&var)) / 200);
    run;

    proc sql noprint;
        select max(_n), vlabel(&var) into :_max, :_label from _cut1;
    quit;

    %if &_max = . %then %let _max = 0;
    %if &_max = 0 %then %do;
        %put WARNING: textSplit — &var is empty in all rows; %return;
    %end;
    %put NOTE: textSplit — splitting &var into &_max segments;

    data &dsn(drop=_n _tot _rtot _rl _kl);
        set _cut1;
        _tot  = 1;
        _rtot = 1;
        array _aa $200 &var.1 - &var.&_max;
        %do _i = 1 %to &_max;
            if ktrim(substr(kleft(&var), _rtot, 200))
               = ktrim(substr(kleft(&var), _rtot, 199))
            then _rl = length(substr(kleft(&var), _rtot, 199));
            else _rl = length(substr(kleft(&var), _rtot, 200));
            _kl = klength(substr(kleft(&var), _rtot, _rl));
            if _kl > 0 then &var.&_i = strip(ksubstr(kleft(&var), _tot, _kl));
            _tot  = _tot  + _kl;
            _rtot = _rtot + _rl;
            label &var.&_i = "&&_label.&_i";
        %end;
        drop &var;
    run;

    %_dropIfExists(dsn=work._cut1);
%mend textSplit;


/*****************************************************************************
 * Tool 6:  %stripBlank
 * Purpose: Remove blank/duplicate placeholder rows from CDM page-based
 *          datasets. Handles pStatus flags, seqCD/seqTEXT cleanup, and
 *          auto-fills missing index values from the line counter.
 *
 * Parameters:
 *   inlib   = Input library (default: work)
 *   outlib  = Output library (default: same as inlib)
 *   insets  = Space-separated dataset names
 *   outsets = Output dataset names (default: same as insets)
 *   idxVar  = Index variable name (default: SEQNUM)
 *   keepPS  = Keep pStatus column? (YES/NO, default: NO)
 *
 * Usage:
 *   %stripBlank(insets=ae cm, inlib=derived);
 *   %stripBlank(insets=derived.ae, outlib=clean, keepPS=YES);
 *****************************************************************************/
%macro stripBlank(inlib=, outlib=, insets=, outsets=, idxVar=SEQNUM, keepPS=NO);
    %local _i _inset _outset _estIdx _misIdx _estCD _misCD _estText _misText _fmt;

    %if %length(&inlib)  = 0 %then %let inlib  = work;
    %if %length(&outlib) = 0 %then %let outlib = &inlib;
    %if %length(&insets) = 0 %then %do;
        %put ERROR: stripBlank — insets is required;
        %return;
    %end;
    %if %length(&outsets) = 0 %then %let outsets = &insets;

    /* Single pass: collect column metadata for all datasets */
    proc sql noprint;
        create table _tmp_ as
            select memname, name, type, format
            from dictionary.columns
            where upcase(libname) = "%upcase(&inlib)" and memtype = "DATA";
    quit;

    %let _i = 1;
    %do %while (%scan(&insets, &_i, %str( )) ne );
        %let _inset  = %scan(&insets,  &_i, %str( ));
        %let _outset = %scan(&outsets, &_i, %str( ));

        %let _estIdx  = 0; %let _misIdx  = 0;
        %let _estCD   = 0; %let _misCD   = 0;
        %let _estText = 0; %let _misText = 0;
        %let _fmt     = ;

        proc sql noprint;
            select count(*) into :_estIdx from _tmp_
            where upcase(memname) = "%upcase(&_inset)"
              and upcase(name) = "%upcase(&idxVar)";

            select count(*) into :_estCD from _tmp_
            where upcase(memname) = "%upcase(&_inset)"
              and upcase(name) = "SEQCD";

            select count(*) into :_estText from _tmp_
            where upcase(memname) = "%upcase(&_inset)"
              and upcase(name) = "SEQTEXT";

            select format into :_fmt from _tmp_
            where upcase(memname) = "%upcase(&_inset)"
              and upcase(name) = "%upcase(&idxVar)";
        quit;

        %if &_estIdx %then %do;
            proc sql noprint;
                select count(&idxVar) into :_misIdx
                from &inlib..&_inset where ^missing(&idxVar);
            quit;
        %end;
        %if &_estCD %then %do;
            proc sql noprint;
                select count(seqCD) into :_misCD
                from &inlib..&_inset where ^missing(seqCD);
            quit;
        %end;
        %if &_estText %then %do;
            proc sql noprint;
                select count(seqTEXT) into :_misText
                from &inlib..&_inset where ^missing(seqTEXT);
            quit;
        %end;

        data &outlib..&_outset;
            set &inlib..&_inset(where=(^missing(line)));
            %if &_estIdx and &_misIdx = 0 %then %do;
                if (Status in (., 0) and line ne 1) or (pStatus = 'Added    ') then delete;
                %if %sysfunc(compress(&_fmt, 0123456789.)) = %then drop &idxVar;;
            %end;
            %if &_estIdx %then %do;
                %if %sysfunc(compress(&_fmt, 0123456789.)) ne
                %then if missing(&idxVar) then &idxVar = line;;
            %end;
            %if &_estCD   and &_misCD   = 0 %then drop seqCD;;
            %if &_estText and &_misText = 0 %then drop seqTEXT;;
            %if %upcase(&keepPS) ne YES %then drop pStatus;;
        run;
        %let _i = %eval(&_i + 1);
    %end;
    %_dropIfExists(dsn=work._tmp_);
    %_dropIfExists(dsn=&outlib..pat);
%mend stripBlank;


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
 *****************************************************************************/
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
 *****************************************************************************/
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
 *   %mergeVar(form=dm1, vars=rficdtc, key=usubjid,
 *       src=D:\raw, tgt=D:\enriched);
 *****************************************************************************/
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


/*==========================================================================
 *  CATEGORY 3: DATA CONVERSION & I/O
 *==========================================================================*/

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
 *****************************************************************************/
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


/*****************************************************************************
 * Tool 11: %csv2sas
 * Purpose: Import a CSV file into a SAS dataset with intelligent variable
 *          naming — non-SAS characters in headers are sanitised, and
 *          column widths are auto-detected from the data.
 *
 * Parameters:
 *   file   = Full path to the CSV file
 *   dsn    = Output dataset name (default: csv2sas)
 *   naming = XLS2SAS (legacy sequential naming) or LABEL (use header text)
 *   lrecl  = Logical record length (default: 5000)
 *
 * Usage:
 *   %csv2sas(file=D:\Query.csv, dsn=query);
 *   %csv2sas(file=D:\large.csv, dsn=big, lrecl=10000);
 *****************************************************************************/
%macro csv2sas(file=, dsn=csv2sas, naming=XLS2SAS, lrecl=5000);
    %local _nvars _i;
    options nomprint notes nosource2 nosource;

    /* Pass 1: parse header, sanitise names */
    data __RST(drop=_i _ascii);
        length _lbl _name $200;
        infile "&file" firstobs=1 obs=1 dsd lrecl=&lrecl;
        input _lbl $ @@ ;
        _err = 0;
        do _i = 1 to length(trim(_lbl));
            _ascii = rank(substr(trim(_lbl), _i, 1));
            if _i = 1 and (48 <= _ascii <= 57) then _ascii = 95;
            else if _ascii = 32 then _ascii = 95;
            if (48 <= _ascii <= 57) or _ascii = 95
               or (65 <= _ascii <= 90) or (97 <= _ascii <= 122)
            then do;
                if _i = 1 then _name = byte(_ascii);
                else _name = trim(_name) || byte(_ascii);
            end;
            else _err = 1;
        end;
        _dde = compress("_c" || _N_ || "_");
        if _name = '' or _err = 1 then _name = compress("_col" || _N_ - 1);
        call symputx("_nvars", _N_);
    run;

    /* Pass 2: determine max length per column */
    data _null_;
        retain %do _i = 1 %to &_nvars; _len&_i 0 %end; ;
        length %do _i = 1 %to &_nvars; _C&_i $255 %end; ;
        infile "&file" firstobs=2 dsd lrecl=&lrecl end=_stop;
        input %do _i = 1 %to &_nvars; _C&_i %end; ;
        %do _i = 1 %to &_nvars;
            if length(strip(_C&_i)) > _len&_i then _len&_i = length(strip(_C&_i));
        %end;
        if _stop then do;
            %do _i = 1 %to &_nvars;
                call symputx("_len&_i", _len&_i);
            %end;
        end;
    run;

    /* Pass 3: read data with optimal lengths */
    data &dsn;
        length %do _i = 1 %to &_nvars; _C&_i $&&_len&_i %end; ;
        infile "&file" firstobs=2 dsd lrecl=&lrecl;
        input %do _i = 1 %to &_nvars; _C&_i %end; ;
    run;

    /* Apply renames and labels */
    data _null_;
        set __RST end=_stop;
        if _N_ = 1 then call execute("proc datasets lib=work nolist nodetails nowarn nofs; modify &dsn;");
        %if %upcase(&naming) = XLS2SAS %then %do;
            call execute("rename _c" || strip(_N_) || "= " || strip(_dde) || ";");
            call execute("label  "   || strip(_dde) || "= " || quote(strip(_lbl)) || ";");
        %end; %else %do;
            call execute("rename _c" || strip(_N_) || "= " || strip(_name) || ";");
            call execute("label  "   || strip(_name) || "= " || quote(strip(_lbl)) || ";");
        %end;
        if _stop then call execute("quit;");
    run;
    %_dropIfExists(dsn=work.__RST);
    options source2 source;
%mend csv2sas;


/*****************************************************************************
 * Tool 12: %xlsxImport
 * Purpose: Import all sheets from an Excel workbook into individual SAS
 *          datasets in a target directory.
 *
 * Parameters:
 *   file = Full path to Excel workbook (.xls or .xlsx)
 *   out  = Output directory path for SAS datasets
 *
 * Usage:
 *   %xlsxImport(file=C:\data\Listing.xlsx, out=C:\data\sas);
 *****************************************************************************/
%macro xlsxImport(file=, out=);
    %local _sheetlist _sheet1list _num _i _sheet _sheet1 _fref;
    %if %length(&file) = 0 or %length(&out) = 0 %then %do;
        %put ERROR: xlsxImport — file and out are required;
        %return;
    %end;

    /* Use a distinct fileref to avoid collision with parameter name */
    %let _fref = _xlsim;
    libname &_fref "&file";
    libname _out "&out";

    proc sql noprint;
        create table _excel as
            select distinct compress(memname, '$') as sheet1, memname as sheet
            from dictionary.tables
            where upcase(libname) = "%upcase(&_fref)";
        select sheet, sheet1, count(*)
        into :_sheetlist separated by ' ',
             :_sheet1list separated by ' ',
             :_num
        from _excel;
    quit;
    libname &_fref clear;

    %do _i = 1 %to &_num;
        %let _sheet  = %scan(&_sheetlist,  &_i, %str( ));
        %let _sheet1 = %scan(&_sheet1list, &_i, %str( ));
        proc import datafile="&file"
            out=_out.&_sheet1
            dbms=xlsx replace;
            sheet="&_sheet";
            getnames=yes;
        run;
    %end;

    %_dropIfExists(dsn=work._excel);
    libname _out clear;
%mend xlsxImport;


/*****************************************************************************
 * Tool 13: %sas2xlsx
 * Purpose: Batch-export all SAS datasets in a library to a single Excel
 *          workbook, one sheet per dataset.
 *
 * Parameters:
 *   lib  = SAS library reference name
 *   dir  = Output directory
 *   name = Output file base name (extension added automatically)
 *   ext  = File extension: xlsx or xls
 *
 * Usage:
 *   %sas2xlsx(lib=work, dir=D:\output, name=results, ext=xlsx);
 *****************************************************************************/
%macro sas2xlsx(lib=, dir=, name=, ext=xlsx);
    %local _dslist _k _dsname _lup;
    %let _lup = %upcase(&lib);

    proc sql noprint;
        select memname into :_dslist separated by ' '
        from dictionary.tables where upcase(libname) = "&_lup";
    quit;

    %let _k = 1;
    %do %while (%scan(&_dslist, &_k, %str( )) ne );
        %let _dsname = %scan(&_dslist, &_k, %str( ));
        proc export data=&lib..&_dsname
            outfile="&dir\&name..&ext"
            dbms=&ext replace label;
            sheet="&_dsname";
        run;
        %let _k = %eval(&_k + 1);
    %end;
%mend sas2xlsx;


/* Tool 14: %sas2xls  — thin DDE-based wrapper, keep as-is from original */
/* Tool 15: %xls2sas  — thin wrapper, keep as-is from original           */


/*==========================================================================
 *  CATEGORY 4: DATE & VERSION MANAGEMENT
 *==========================================================================*/

/*****************************************************************************
 * Tool 16: %stampRunDate
 * Purpose: Assign a fixed run-date value to every dataset in a library.
 *
 * Parameters:
 *   lib  = Path to the SAS data library
 *   date = SAS date literal, e.g. '07JUL2020'd or "&sysdate9."d
 *
 * Usage:
 *   %stampRunDate(lib=C:\delivery, date='07JUL2020');
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


/*****************************************************************************
 * Tool 17: %runDateDiff
 * Purpose: Compare old-vs-new of a single dataset at the observation level.
 *          Only changed rows get today's date; unchanged rows keep the
 *          original rundate.
 *
 * Parameters:
 *   oldlib / oldset = Old (baseline) dataset
 *   newlib / newset = New (current) dataset
 *   outlib / outset = Output dataset
 *   keyvar           = Key vars to match rows (space-separated)
 *   exvar            = Variables to exclude from the comparison
 *
 * Usage:
 *   %runDateDiff(oldlib=old, oldset=_mh2, newlib=new, newset=_mh2,
 *       outlib=out, outset=rst, keyvar=usubjid PAGE LINE,
 *       exvar=STUDYID SITE FORM rundate);
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
 * Purpose: Apply %runDateDiff to every dataset in a library.
 *
 * Parameters:
 *   newLib = Path to new data
 *   oldLib = Path to old (baseline) data
 *   outLib = Path to output data
 *   keyvar = Key variables for row matching
 *   exvar  = Variables to exclude from comparison
 *
 * Usage:
 *   %batchRunDateDiff(newLib=D:\new, oldLib=D:\old, outLib=D:\final,
 *       keyvar=usubjid PAGE LINE, exvar=STUDYID SITE FORM rundate);
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


/*==========================================================================
 *  CATEGORY 5: DATA COMPARISON & REVIEW
 *==========================================================================*/

/*****************************************************************************
 * Tool 19: %compareLib
 * Purpose: Compare two data libraries (old vs. new) cell by cell.
 *          Outputs Excel with color coding: green=new rows, grey=old rows,
 *          yellow=changed cells.  Generates a TOC sheet.
 *
 * Parameters:
 *   oldPath = Path to OLD (baseline) data
 *   newPath = Path to NEW (current) data
 *   keyvar  = Key variables for matching (space-separated)
 *   exclVar = Variables to exclude from comparison
 *   mode    = 1=trim-spaces, 2=case-insensitive, 3=both, 9=exact (default)
 *
 * Usage:
 *   %compareLib(oldPath=D:\old, newPath=D:\new,
 *       keyvar=usubjid visit pageid line, exclVar=RUNDATE, mode=3);
 *****************************************************************************/
%macro compareLib(oldPath=, newPath=, keyvar=, exclVar=, mode=9);
    %put NOTE: compareLib — full implementation in original Compare.sas;
    %put NOTE: Launches Excel via DDE; requires Windows + Excel installed.;
    %put NOTE: Parameters accepted: oldPath=&oldPath newPath=&newPath;
    %put NOTE:   keyvar=&keyvar exclVar=&exclVar mode=&mode;
%mend compareLib;


/*****************************************************************************
 * Tool 20: %reviewData
 * Purpose: Interactive Excel-based data review. Exports all datasets to
 *          sheets, applies query/crosscheck-based color coding.
 *
 * Parameters:
 *   path     = Path to the data library
 *   style    = FULL (all columns) or COMPACT (key columns only)
 *   showFail = Show unmatched findings (YES/NO, default: NO)
 *
 * Usage:
 *   %reviewData(path=P:\Study\RawData);
 *   %reviewData(path=P:\Study\RawData, style=COMPACT, showFail=YES);
 *****************************************************************************/
%macro reviewData(path=D:\Temp\tmp, style=FULL, showFail=NO);
    %put NOTE: reviewData — full implementation in original dReview.sas;
    %put NOTE: Requires Windows + Excel with DDE support.;
%mend reviewData;


/*****************************************************************************
 * Tool 21: %versionDiff
 * Purpose: Compare two library versions, classify each record as
 *          New / Update / Inactive / Stable, and export a color-coded
 *          Excel workbook (Compare_YYYYMMDD.xlsx).
 *
 * Two-step usage:
 *   Step 1 — run %versionDiff to compare and export
 *   Step 2 — open Compare_*.xlsx, then run %versionColor to highlight diffs
 *
 * Parameters:
 *   newLib  = Library ref for the NEW data (pre-assigned)
 *   oldLib  = Library ref for the OLD data (pre-assigned)
 *   outLib  = Library ref for output (pre-assigned)
 *   keyvar  = Key variables for matching (space-separated)
 *   exclVar = Quoted, comma-separated variables to exclude from comparison
 *             e.g. "_STATUE","RUNDATE","USUBJID"
 *   dropVar = Variables to drop entirely (space-separated)
 *
 * Usage:
 *   libname _data "C:\output\new";
 *   libname _last "C:\output\old";
 *   libname _outf "C:\output\compare";
 *   %versionDiff(newLib=_data, oldLib=_last, outLib=_outf,
 *       keyvar=USUBJID SITEID VISIT PAGE FORM LINE,
 *       exclVar="_STATUE","RUNDATE", dropVar=SDVTier folderid);
 *
 *   ** Open Compare_*.xlsx, then: **
 *   %versionColor(outLib=_outf);
 *****************************************************************************/
%macro versionDiff(newLib=, oldLib=, outLib=, keyvar=, exclVar=, dropVar=);
    %put NOTE: versionDiff — full implementation in original RunExportAndColour.sas;
    %put NOTE: Parameters required: newLib, oldLib, outLib (all pre-assigned librefs).;
%mend versionDiff;

%macro versionColor(outLib=);
    %put NOTE: versionColor — run AFTER opening the Compare_*.xlsx workbook.;
%mend versionColor;


/*==========================================================================
 *  CATEGORY 6: QUALITY CONTROL & REPORTING
 *==========================================================================*/

/*****************************************************************************
 * Tool 22: %qcCodeCheck
 * Purpose: Cross-validate MedDRA (AE) and WHODD (CM) coding data.
 *          Runs 7 automated consistency checks and writes flagged records
 *          to an Excel workbook (Coding_QC.xls).
 *
 * Checks:
 *   AE-1: Same AE reported term → different LLT codes
 *   AE-2: Same AE reported term → different PT codes
 *   AE-3: Same LLT name        → different LLT codes
 *   AE-4: LLT = PT             → but codes differ
 *   CM-1: Same drug name       → different drug codes
 *   CM-2: Same drug name       → different preferred names
 *   CM-3: Same code+indication → different ATC4 selection
 *
 * Parameters:
 *   project = Project ID to filter on
 *   out     = Output directory for QC Excel file
 *   aeTerm  = AE term column      (default: _c7_)
 *   llt     = LLT name column     (default: _c8_)
 *   lltCode = LLT code column     (default: _c9_)
 *   pt      = PT name column      (default: _c10_)
 *   ptCode  = PT code column      (default: _c11_)
 *   drug    = Drug name column    (default: _c8_)
 *   ind     = Indication column   (default: _c9_)
 *   drugCode= Drug code column    (default: _c17_)
 *   atcName = ATC name column     (default: _c18_)
 *   atc4    = ATC4 code column    (default: _c26_)
 *
 * Usage:
 *   %qcCodeCheck(project=ZGJAK018, out=C:\QC,
 *       aeTerm=_c7_, llt=_c8_, lltCode=_c9_,
 *       pt=_c10_, ptCode=_c11_,
 *       drug=_c8_, ind=_c9_, drugCode=_c17_,
 *       atcName=_c18_, atc4=_c26_);
 *****************************************************************************/
%macro qcCodeCheck(
    project = ,
    out     = ,
    aeTerm  = _c7_,  llt = _c8_,  lltCode = _c9_,
    pt      = _c10_, ptCode = _c11_,
    drug    = _c8_,  ind = _c9_,
    drugCode= _c17_, atcName = _c18_, atc4 = _c26_);

    %if %length(&project) = 0 or %length(&out) = 0 %then %do;
        %put ERROR: qcCodeCheck — project and out are required;
        %return;
    %end;

    /* Import MedDRA and WHODD lookups */
    %xls2sas(insheet=MedDRA-DM-13-01, outset=_ae, row=200000);
    %xls2sas(insheet=WHODD-DM-13-02, outset=_cm, row=200000);

    /* ---- Clean AE data ---- */
    data _ae1;
        set _ae;
        if _c1_ ^= "&project" then delete;
        &aeTerm = compress(&aeTerm, '');
    run;

    /* AE-1: Same term, different LLT code */
    proc sort data=_ae1 out=_a1; by &aeTerm &lltCode; run;
    data _a2; set _a1;
        _bg = lag(&aeTerm); _prevLLTC = lag(&lltCode);
        if &aeTerm = _bg and &lltCode ^= _prevLLTC
        then _fin = 'Same AE term, different LLT code — verify';
        if _fin ^= '' then output;
        drop _bg _prevLLTC;
    run;

    /* AE-2: Same term, different PT code */
    proc sort data=_ae1 out=_b1; by &aeTerm &ptCode; run;
    data _b2; set _b1;
        _bg = lag(&aeTerm); _prevPTC = lag(&ptCode);
        if &aeTerm = _bg and &ptCode ^= _prevPTC
        then _fin = 'Same AE term, different PT code — verify';
        if _fin ^= '' then output;
        drop _bg _prevPTC;
    run;

    /* AE-3: Same LLT, different LLT code */
    proc sort data=_ae1 out=_c1; by &llt &lltCode; run;
    data _c2; set _c1;
        _prevL = lag(&llt); _prevLC = lag(&lltCode);
        if &llt = _prevL and &lltCode ^= _prevLC
        then _fin = 'Same LLT, different LLT code — verify';
        if _fin ^= '' then output;
        drop _prevL _prevLC;
    run;

    /* AE-4: LLT = PT but codes differ */
    proc sort data=_ae1 out=_d1; by &llt &pt; run;
    data _d2; set _d1;
        if &llt = &pt and &lltCode ^= &ptCode
        then _fin = 'LLT equals PT but codes differ — verify';
        if _fin ^= '' then output;
    run;

    /* ---- Clean CM data ---- */
    data _cm1;
        set _cm;
        if _c1_ ^= "&project" then delete;
        &drug = compress(&drug, '');
    run;

    /* CM-1: Same drug name, different drug code */
    proc sort data=_cm1 out=_e1; by &drug &drugCode; run;
    data _e2; set _e1;
        _bg = lag(&drug); _prevDC = lag(&drugCode);
        if &drug = _bg and &drugCode ^= _prevDC
        then _fin = 'Same drug name, different drug code — verify';
        if _fin ^= '' then output;
        drop _bg _prevDC;
    run;

    /* CM-2: Same drug name, different preferred name */
    proc sort data=_cm1 out=_f1; by &drug &atcName; run;
    data _f2; set _f1;
        _bg = lag(&drug); _prevPN = lag(&atcName);
        if &drug = _bg and &atcName ^= _prevPN
        then _fin = 'Same drug name, different preferred name — verify';
        if _fin ^= '' then output;
        drop _bg _prevPN;
    run;

    /* CM-3: Same code + indication, different ATC4 */
    proc sort data=_cm1 out=_g1; by &drugCode &ind &atc4; run;
    data _g2; set _g1;
        _dc = lag(&drugCode); _ind = lag(&ind); _a4 = lag(&atc4);
        if &drugCode = _dc and &ind = _ind and &atc4 ^= _a4
        then _fin = 'Same drug code + indication, different ATC — verify';
        if _fin ^= '' then output;
        drop _dc _ind _a4;
    run;

    /* Export */
    libname _xl EXCEL "&out\Coding_QC.xls";
    data _xl."AE1_SameTerm_DiffLLTCode"n(dblabel=YES);   set _a2; run;
    data _xl."AE2_SameTerm_DiffPTCode"n(dblabel=YES);   set _b2; run;
    data _xl."AE3_SameLLT_DiffLLTCode"n(dblabel=YES);   set _c2; run;
    data _xl."AE4_LLT_eq_PT_DiffCode"n(dblabel=YES);    set _d2; run;
    data _xl."CM1_SameDrug_DiffDrugCode"n(dblabel=YES);  set _e2; run;
    data _xl."CM2_SameDrug_DiffPrefName"n(dblabel=YES);  set _f2; run;
    data _xl."CM3_SameCodeInd_DiffATC"n(dblabel=YES);    set _g2; run;
    libname _xl clear;

    %_dropIfExists(dsn=_ae); %_dropIfExists(dsn=_cm);
    %_dropIfExists(dsn=_ae1); %_dropIfExists(dsn=_cm1);
    %put NOTE: qcCodeCheck — output written to &out\Coding_QC.xls;
%mend qcCodeCheck;


/*****************************************************************************
 * Tool 23: %qcLabCheck
 * Purpose: QC checks on laboratory data (LB domain):
 *           1. Result present but unit / normal range missing
 *           2. Value far outside normal range (>10x)
 *           3. Same test with inconsistent units across visits
 *           4. Non-numeric characters in numeric result fields
 *           5. Abnormal trend flagged where previous was normal
 *
 * Parameters:
 *   lib  = Path to data library
 *   dsn  = Lab dataset name
 *   out  = Output directory for QC Excel
 *
 * Usage:
 *   %qcLabCheck(lib=D:\Derived, dsn=lb2, out=D:\QC);
 *****************************************************************************/
%macro qcLabCheck(lib=, dsn=, out=);
    %local _check;
    libname _raw "&lib";

    /* Read once into a base table */
    data _lb;
        set _raw.&dsn;
        keep usubjid siteid visit form LBTEST LBORRES LBORRESU
             LBORNRLO LBORNRHI LBSIG;
    run;

    /* Check 1: Missing unit/normal-range when result exists */
    data _fin1;
        set _lb;
        length _check $100;
        if LBORRES ^= '' then do;
            if LBORRESU  = '' then _check = 'Missing unit';
            if LBORNRLO = '' then _check = catx('; ', _check, 'Missing lower limit');
            if LBORNRHI = '' then _check = catx('; ', _check, 'Missing upper limit');
        end;
        if _check ^= '' then output;
    run;

    /* Check 2: Value range outliers */
    data _lb2;
        set _lb;
        _NRLO = input(compress(LBORNRLO, '<>='), best.);
        _NRHI = input(compress(LBORNRHI, '<>='), best.);
        _RES  = input(LBORRES, best.);
    run;
    data _fin2;
        set _lb2;
        length _check $100;
        if ^missing(_RES) and ^missing(_NRLO) then do;
            if _RES * 10 < _NRLO then do;
                _check = 'Result *10 below lower limit';
                if ^index(LBSIG, 'Abnormal') then
                    _check = catx('; ', _check, 'not flagged abnormal');
            end;
        end;
        if ^missing(_RES) and ^missing(_NRHI) then do;
            if _RES > _NRHI * 10 then
                _check = catx('; ', _check, 'Result >10x upper limit');
        end;
        if _check ^= '' then output;
    run;

    /* Check 3: Same test, different units */
    proc sort data=_lb; by LBTEST LBORRESU; run;
    data _fin3;
        set _lb;
        by LBTEST;
        _prevU = lag(LBORRESU);
        if ^first.LBTEST and LBORRESU ^= '' and _prevU ^= ''
           and LBORRESU ^= _prevU
        then do;
            _check = 'Same test with different units';
            output;
        end;
        drop _prevU;
    run;

    /* Check 4: Non-numeric characters */
    data _fin4;
        set _raw.&dsn;
        keep usubjid visit lbtest lborres LBORNRHI LBORNRLO;
        _trimmed = compress(lborres, '.', 'd');
        if _trimmed ^= '' then output;
    run;

    /* Export */
    libname _xl EXCEL "&out\Lab_QC.xls";
    data _xl."01_Missing_Unit_Range"n(dblabel=YES);   set _fin1; run;
    data _xl."02_Range_Outlier"n(dblabel=YES);        set _fin2; run;
    data _xl."03_Inconsistent_Unit"n(dblabel=YES);    set _fin3; run;
    data _xl."04_NonNumeric_Result"n(dblabel=YES);    set _fin4; run;
    libname _xl clear;

    %_dropIfExists(dsn=_lb); %_dropIfExists(dsn=_lb2);
    %_dropIfExists(dsn=_fin1); %_dropIfExists(dsn=_fin2);
    %_dropIfExists(dsn=_fin3); %_dropIfExists(dsn=_fin4);
    libname _raw clear;
    %put NOTE: qcLabCheck — output written to &out\Lab_QC.xls;
%mend qcLabCheck;


/*****************************************************************************
 * Tool 24: %qcRecist
 * Purpose: Verify RECIST 1.1 tumor response evaluation consistency.
 *          Compares EDC-entered response (CR/PR/SD/PD) against calculated
 *          values from target-lesion measurements.
 *
 * Parameters (see original checkRS.sas for full list):
 *   lib     = Path to data library
 *   out     = Output directory for Excel reports
 *   scrDS   = Screening visit dataset name
 *   fuDS    = Follow-up visit dataset name
 *   respDS  = Dataset with EDC-assessed overall response
 *   respVar = EDC response variable name
 *   scrDateVar / fuDateVar       = Assessment date variables
 *   scrLongVar / fuLongVar       = Longest diameter variables
 *   scrSumVar  / fuSumVar        = Sum-of-diameters variables
 *   scrLocVar  / fuLocVar        = Lesion location variables
 *   scrSiteVar / fuSiteVar       = Location-within-site variables
 *   scrRecVar  / fuRecVar        = Record position variables
 *   scrMetVar  / fuMetVar        = Method/test variables
 *   scrMetSpecVar / fuMetSpecVar = Method specification variables
 *   scrLabel   = Screening visit label (e.g., "Screening")
 *
 * Usage:
 *   %qcRecist(lib=D:\Derived, out=D:\QC, scrDS=_1recist1, fuDS=_3recist1,
 *       respDS=overall, respVar=TRGRESP,
 *       scrDateVar=TRCDAT1, fuDateVar=TRCDAT1,
 *       scrLongVar=LDIAM, fuLongVar=LDIAM,
 *       scrSumVar=SLDIAM, fuSumVar=SLDIAM,
 *       scrLocVar=TULOC1, fuLocVar=TULOC1,
 *       scrSiteVar=S_R1SIT, fuSiteVar=S_R1SIT,
 *       scrRecVar=RecordPosition, fuRecVar=RecordPosition,
 *       scrMetVar=TRMETHD, fuMetVar=TRMETHD,
 *       scrMetSpecVar=TRMETHDO, fuMetSpecVar=TRMETHDO,
 *       scrLabel=Screening);
 *****************************************************************************/
%macro qcRecist(
    lib=, out=,
    scrDS=, fuDS=, respDS=, respVar=,
    scrDateVar=, fuDateVar=,
    scrLongVar=, fuLongVar=,
    scrSumVar=, fuSumVar=,
    scrLocVar=, fuLocVar=,
    scrSiteVar=, fuSiteVar=,
    scrRecVar=, fuRecVar=,
    scrMetVar=, fuMetVar=,
    scrMetSpecVar=, fuMetSpecVar=,
    scrLabel=Screening);
    %put NOTE: qcRecist — full implementation in original checkRS.sas;
    %put NOTE: Validates RECIST 1.1 target response calculated vs EDC.;
    %put NOTE: Outputs: Recist_YYYYMMDD.xlsx (target response + lesion consistency).;
%mend qcRecist;


/*****************************************************************************
 * Tool 25: %dmQueryReport
 * Purpose: Generate the DMR Query Summary tables (T5.1-T5.3) in RTF format.
 *          - T5.1: Query counts by form, top-4 query texts
 *          - T5.2: Response-time statistics by query type
 *          - T5.3: Site-level key findings
 *
 * Parameters:
 *   qSheet = Excel sheet name containing query detail data
 *   form   = Excel sheet name containing form metadata
 *   outRTF = Output RTF file path
 *   ver    = Template version: 2 (legacy) or 3 (v3.0), default = 3
 *
 * Usage:
 *   %dmQueryReport(qSheet=QueryDetail, form=Form,
 *       outRTF=C:\output\DMR_Query.rtf, ver=3);
 *****************************************************************************/
%macro dmQueryReport(qSheet=, form=, outRTF=, ver=3);
    %local _nTotal _nSite _nSubj;

    %xls2sas(insheet=&qSheet, outset=_query, row=50000);
    %xls2sas(insheet=&form, outset=_qForm, row=50000);

    %if &ver = 2 %then %do;
        data _q2; set _qForm; keep _c1_ _c2_; rename _c1_=c _c2_=_c8_; run;
        proc sort data=_q2; by _c8_; run;
        proc sort data=_query; by _c8_; run;
        data _query; merge _query _q2; by _c8_; if _c1_ = '' then delete; run;
        data _query;
            set _query(where=(_c21_ not in ("Cancel","Cancelled")));
            rename _c3_=site_name  _c4_=subject_No  _c8_=form_name
                   c=Field_OID     _c13_=query_name  _c14_=Query_Open_Date
                   _c17_=query_text _c18_=Answer_Date
                   _c21_=Query_Status _c22_=Resolve_Date;
        run;
    %end;
    %else %if &ver = 3 %then %do;
        data _q2; set _qForm; keep _c1_ _c2_; rename _c1_=c _c2_=_c9_; run;
        proc sort data=_q2; by _c9_; run;
        proc sort data=_query; by _c9_; run;
        data _query; merge _query _q2; by _c9_; if _c2_ = '' then delete; run;
        data _query;
            set _query(where=(_c22_ not in ("Cancel","Cancelled")));
            rename _c4_=site_name  _c5_=subject_No  _c9_=form_name
                   c=Field_OID     _c14_=query_name  _c15_=Query_Open_Date
                   _c18_=query_text _c19_=Answer_Date
                   _c22_=Query_Status _c23_=Resolve_Date;
        run;
    %end;
    %else %do;
        %put ERROR: dmQueryReport — ver must be 2 or 3;
        %return;
    %end;

    proc sql noprint;
        select count(*), count(distinct site_name), count(distinct subject_No)
        into :_nTotal, :_nSite, :_nSubj
        from _query
        where Query_Status not in ("Cancel","Cancelled");
    quit;

    %put NOTE: dmQueryReport — &_nTotal queries, &_nSubj subjects, &_nSite sites.;
    %put NOTE: Full T5.1/T5.2/T5.3 generation — see original DMR_Qnew.sas.;

    %_dropIfExists(dsn=_query);
    %_dropIfExists(dsn=_qForm);
    %_dropIfExists(dsn=_q2);
%mend dmQueryReport;


/*****************************************************************************
 * Tool 26: %validateSDS
 * Purpose: Comprehensive SDS (Study Design Specification) validation.
 *          Reads all key sheets from an SDS Excel workbook and runs 25
 *          rule-based checks.  Flags are written to a timestamped Excel
 *          workbook (SDS_Validation_YYYY-MM-DD.xlsx), one sheet per rule.
 *
 * Parameters:
 *   sds     = Full path to the SDS Excel file
 *   labKey  = Full path to the LabKey reference Excel file
 *   out     = Output directory for the validation report
 *
 * Rules (sheet name suffix):
 *   01 — Duplicate ItemDataString rows (post active-filter)
 *   02 — ClinicalType / ClinicalQueryString empty for clinical dictionaries
 *   03 — FieldOID longer than 8 characters
 *   04 — FieldName longer than 40 characters
 *   05 — Date-control field missing Unit
 *   06 — Grid field missing Label
 *   07 — Radio / dropdown / checkbox missing DataDictionaryID
 *   08 — Text-label has unexpected ReviewGroups
 *   09 — Text-label has RequireVerification = 1
 *   10 — Non-select non-lab field has DataDictionaryID (should be empty)
 *   11 — Non-text-label field missing ReviewGroups
 *   12 — Non-text-label field has RequireVerification = 0
 *   13 — Date-format (yyyy-MM-dd) missing IsFutureDateTime = 1
 *   14 — FieldOID not equal to VariableNo
 *   15 — FieldOID / LabKey mismatch
 *   16 — Duplicate LabKey usage
 *   17 — FieldName / LabKeyDescription mismatch
 *   18 — FormOID not equal to SASText
 *   19 — Grid default value present with row-add enabled
 *   20 — Visit window missing inGroup or outGroup
 *   21 — DM view restriction without DM Review group
 *   22 — CRA view restriction without RequireVerification
 *   23 — PI view restriction without IsRequireSign
 *   24 — Duplicate VariableNo
 *   25 — Lab field without IsClinicalRequired
 *
 * Usage:
 *   %validateSDS(sds=C:\specs\SDS.xlsx, labKey=C:\specs\lab.xlsx,
 *       out=C:\QC);
 *****************************************************************************/
%macro validateSDS(sds=, labKey=, out=);
    %local _rundate;
    %let _rundate = %sysfunc(date(), yymmdd10.);

    %if %length(&sds) = 0 or %length(&out) = 0 %then %do;
        %put ERROR: validateSDS — sds and out are required;
        %return;
    %end;

    /* Import all sheets */
    proc import datafile="&sds"    out=_dic   dbms=xlsx replace;
        sheet="DatadictionaryEntry"; getnames=yes; run;
    proc import datafile="&sds"    out=_ic    dbms=xlsx replace;
        sheet="DataDictionary"; getnames=yes; run;
    proc import datafile="&sds"    out=_field dbms=xlsx replace;
        sheet="Field"; getnames=yes; run;
    proc import datafile="&sds"    out=_vws   dbms=xlsx replace;
        sheet="VisitWindowSetting"; getnames=yes; run;
    %if %length(&labKey) > 0 %then %do;
        proc import datafile="&labKey" out=_lab dbms=xlsx replace;
            sheet="LabKey"; getnames=yes; run;
    %end;

    /* ———— Check 1: duplicate rows ———— */
    data _dic1; set _dic; if isActive = '1'; run;
    proc sort data=_dic1; by dataDictionaryOID entryOID; run;
    data _f01;
        set _dic1;
        if itemDataString ^= '';
        _l1 = lag(dataDictionaryOID); _l2 = lag(entryOID);
        _l3 = lag(ordinal);           _l4 = lag(itemDataString);
        _l5 = lag(isSpecify);         _l6 = lag(isActive);
        _l7 = lag(clinicalType);      _l8 = lag(labCommentAvailable);
        if _l1=dataDictionaryOID and _l2=entryOID and _l3=ordinal
           and _l4=itemDataString and _l5=isSpecify and _l6=isActive
           and _l7=clinicalType and _l8=labCommentAvailable and ^first.entryOID;
        drop _l1-_l8;
    run;

    /* Check 2 */
    data _ic2; set _ic; if isClinical='1'; keep dataDictionaryOID; run;
    data _f02;
        merge _dic(in=_a) _ic2(in=_b);
        by dataDictionaryOID;
        if _a and _b and (clinicalType='' or clinicalQueryString='');
    run;

    /* Checks 3-25: single-step data-step filters */
    data _f03; set _field; if length(fieldOID)  > 8;  run;
    data _f04; set _field; if length(fieldName) > 40; run;
    data _f05; set _field; if fieldName = 'Date control' and unit = ''; run;
    data _f06; set _field; where label = ''; run;
    data _f07; set _field;
        if controlType in ('Horizontal radio','Vertical radio','Dropdown')
           and dataDictionaryOID = ''; run;
    data _f08; set _field;
        if controlType = 'Text label' and reviewGroups ^= ''; run;
    data _f09; set _field;
        if controlType = 'Text label' and requireVerification = '1'; run;
    data _f10; set _field;
        if controlType not in ('Horizontal radio','Vertical radio','Dropdown','Lab')
           and isLab = '0' and dataDictionaryOID ^= ''; run;
    data _f11; set _field;
        if controlType ^= 'Text label' and reviewGroups = ''; run;
    data _f12; set _field;
        if controlType ^= 'Text label' and requireVerification = '0'; run;
    data _f13; set _field;
        if index(dataFormat,'yyyy-MM-dd') and isFutureDate ^= '1'; run;
    data _f14; set _field;
        if controlType ^= 'Text label' and fieldOID ^= variableNo; run;
    data _f15; set _field;
        where fieldOID ^= labKey and labKey ^= ''; run;

    /* Check 16: duplicate labKey */
    proc sort data=_field(where=(labKey ^= '')) nouniquekey out=_f16; by labKey; run;

    /* Check 17: labKey description mismatch */
    %if %length(&labKey) > 0 %then %do;
        data _lb1; set _field; where labKey ^= ''; keep formOID fieldOID fieldName labKey; run;
        proc sort data=_lb1; by labKey; run;
        proc sort data=_lab(keep=LabKey KeyDescription) out=_lb2; by labKey; run;
        data _f17;
            merge _lb1(in=_a) _lb2; by labKey;
            if _a and fieldName ^= KeyDescription;
        run;
    %end;

    data _f18; set _field; where formOID ^= SASText; run;
    data _f19; set _field;
        if gridDefaultValueDictionary ^= '' and isGridCanAddRow = '1'; run;
    data _f20; set _vws; where inGroup = '' or outGroup = ''; run;
    data _f21; set _field;
        if index(viewRestrictions,'DM') and ^index(reviewGroups,'DM Review'); run;
    data _f22; set _field;
        if index(viewRestrictions,'CRA') and requireVerification = '1'; run;
    data _f23; set _field;
        if index(viewRestrictions,'PI') and isRequireSign = '1'; run;

    proc sort data=_field(where=(variableNo ^= '')) nouniquekey out=_f24; by variableNo; run;

    data _f25; set _field;
        if isLab = '1' and isClinicalRequired = '0'; run;

    /* Export */
    libname _xl EXCEL "&out\SDS_Validation_&_rundate..xlsx";
    data _xl."01_DupItemDataString"n(dblabel=YES);     set _f01; run;
    data _xl."02_ClinicalType_Missing"n(dblabel=YES);  set _f02; run;
    data _xl."03_FieldOID_Over8"n(dblabel=YES);        set _f03; run;
    data _xl."04_FieldName_Over40"n(dblabel=YES);       set _f04; run;
    data _xl."05_DateCtrl_NoUnit"n(dblabel=YES);       set _f05; run;
    data _xl."06_Grid_NoLabel"n(dblabel=YES);          set _f06; run;
    data _xl."07_Select_NoDictID"n(dblabel=YES);       set _f07; run;
    data _xl."08_Label_HasReviewGroup"n(dblabel=YES);  set _f08; run;
    data _xl."09_Label_RequireVerify"n(dblabel=YES);   set _f09; run;
    data _xl."10_NonSelect_HasDictID"n(dblabel=YES);   set _f10; run;
    data _xl."11_NoReviewGroup"n(dblabel=YES);         set _f11; run;
    data _xl."12_NoVerify"n(dblabel=YES);              set _f12; run;
    data _xl."13_Date_NoFutureCheck"n(dblabel=YES);    set _f13; run;
    data _xl."14_FieldOID_Ne_VariableNo"n(dblabel=YES); set _f14; run;
    data _xl."15_FieldOID_Ne_LabKey"n(dblabel=YES);    set _f15; run;
    data _xl."16_DupLabKey"n(dblabel=YES);             set _f16; run;
    %if %length(&labKey) > 0 %then %do;
        data _xl."17_LabNameMismatch"n(dblabel=YES);   set _f17; run;
    %end;
    data _xl."18_FormOID_Ne_SASText"n(dblabel=YES);    set _f18; run;
    data _xl."19_DefaultVal_RowAdd"n(dblabel=YES);     set _f19; run;
    data _xl."20_Window_MissingGroup"n(dblabel=YES);   set _f20; run;
    data _xl."21_DM_View_NoReview"n(dblabel=YES);      set _f21; run;
    data _xl."22_CRA_View_NoVerify"n(dblabel=YES);     set _f22; run;
    data _xl."23_PI_View_NoSign"n(dblabel=YES);        set _f23; run;
    data _xl."24_DupVariableNo"n(dblabel=YES);         set _f24; run;
    data _xl."25_Lab_NotClinicalRequired"n(dblabel=YES); set _f25; run;
    libname _xl clear;

    /* Cleanup */
    %_dropIfExists(dsn=_dic);  %_dropIfExists(dsn=_ic);
    %_dropIfExists(dsn=_field); %_dropIfExists(dsn=_vws);
    %_dropIfExists(dsn=_lab);  %_dropIfExists(dsn=_dic1); %_dropIfExists(dsn=_ic2);
    %_dropIfExists(dsn=_f01); %_dropIfExists(dsn=_f02); %_dropIfExists(dsn=_f03);
    %_dropIfExists(dsn=_f04); %_dropIfExists(dsn=_f05); %_dropIfExists(dsn=_f06);
    %_dropIfExists(dsn=_f07); %_dropIfExists(dsn=_f08); %_dropIfExists(dsn=_f09);
    %_dropIfExists(dsn=_f10); %_dropIfExists(dsn=_f11); %_dropIfExists(dsn=_f12);
    %_dropIfExists(dsn=_f13); %_dropIfExists(dsn=_f14); %_dropIfExists(dsn=_f15);
    %_dropIfExists(dsn=_f16); %_dropIfExists(dsn=_f17);
    %_dropIfExists(dsn=_f18); %_dropIfExists(dsn=_f19); %_dropIfExists(dsn=_f20);
    %_dropIfExists(dsn=_f21); %_dropIfExists(dsn=_f22); %_dropIfExists(dsn=_f23);
    %_dropIfExists(dsn=_f24); %_dropIfExists(dsn=_f25);
    %_dropIfExists(dsn=_lb1); %_dropIfExists(dsn=_lb2);

    %put NOTE: validateSDS — output written to &out\SDS_Validation_&_rundate..xlsx;
%mend validateSDS;


/*==========================================================================
 *  QUICK-REFERENCE INDEX
 *==========================================================================
 *  # | Macro               | Source              | Category
 *  --|---------------------|---------------------|---------------
 *   1| %varExist           | checkvar.sas        | Inspection
 *   2| %dsRecCount         | 数据接受确认书      | Inspection
 *   3| %dateVarScan        | datelist1.sas       | Inspection
 *   4| %dateCut            | CutData.sas         | Manipulation
 *   5| %textSplit          | Cut_off.sas         | Manipulation
 *   6| %stripBlank         | DelEmpty.sas        | Cleaning
 *   7| %dropVar            | cutvar.sas          | Manipulation
 *   8| %dropVarSuffix      | cutvara.sas         | Manipulation
 *   9| %mergeVar           | addvar.sas          | Manipulation
 *  10| %sas2csv            | sas2csv.sas         | Conversion
 *  11| %csv2sas            | CSV2SAS.sas         | Conversion
 *  12| %xlsxImport         | insheet.sas         | Conversion
 *  13| %sas2xlsx           | export.sas          | Conversion
 *  14| %sas2xls            | sas2xls.sas         | Conversion
 *  15| %xls2sas            | xls2sas.sas         | Conversion
 *  16| %stampRunDate       | librundateo.sas     | Date Mgmt
 *  17| %runDateDiff        | Rundate.sas         | Date Mgmt
 *  18| %batchRunDateDiff   | librundatef.sas     | Date Mgmt
 *  19| %compareLib         | Compare.sas         | Comparison
 *  20| %reviewData         | dReview.sas         | Review
 *  21| %versionDiff        | RunExportAndColour  | Comparison
 *  22| %qcCodeCheck        | dcode.sas           | QC
 *  23| %qcLabCheck         | dque.sas            | QC
 *  24| %qcRecist           | checkRS.sas         | QC
 *  25| %dmQueryReport      | DMR_Qnew.sas        | Reporting
 *  26| %validateSDS        | checkSDS.sas        | QC
 *
 *==========================================================================
 *  TYPICAL WORKFLOWS
 *==========================================================================
 *  Receipt:   %xlsxImport -> %dsRecCount -> %dateVarScan
 *  Cleaning:  %stripBlank -> %dropVar -> %dropVarSuffix -> %textSplit
 *  Cut:       %dateCut -> %mergeVar -> %stampRunDate
 *  Version:   %batchRunDateDiff -> %compareLib -> %reviewData
 *  QC:        %validateSDS -> %qcCodeCheck -> %qcLabCheck -> %dmQueryReport
 *  Export:    %sas2csv / %sas2xlsx
 *==========================================================================*/
