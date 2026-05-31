/****************************************************************************
 *  Macro: %versionDiff  (+ %versionColor)
 *  Category: Data Comparison
 *  Purpose: Compare two SAS library versions dataset-by-dataset, export
 *           differences to an Excel workbook, and optionally apply cell
 *           colour highlighting via DDE in the open workbook.
 *
 *  %versionDiff  — compares newLib vs oldLib, creates compare_YYYYMMDD.xlsx
 *  %versionColor — highlights changed cells in the already-open Excel workbook
 *                  using DDE (requires Excel with DDE support)
 ****************************************************************************/


/* ========== INTERNAL HELPERS ========== */

%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;

%macro _getDatasetCount(libref=, outvar=_nDs);
    %global &outvar;
    proc sql noprint;
        select count(*) into :&outvar
        from dictionary.tables
        where upcase(libname) = "%upcase(&libref)" and memtype = "DATA";
    quit;
%mend _getDatasetCount;


/* ========== MAIN MACRO: versionDiff ========== */

%macro versionDiff(
    newLib=,        /* NEW libref (pre-assigned) */
    oldLib=,        /* OLD libref (pre-assigned) */
    outLib=,        /* OUTPUT libref (pre-assigned) */
    keyvar=,        /* Key variable(s) for merge, separated by space */
    exclVar=,       /* Variable(s) to exclude from comparison, quoted SAS names */
    dropVar=        /* Variable(s) to drop from datasets */
);

    %local _rundate _dslist _dsetnum _dset _nDs _nowpath _xlsxname _libpath;
    %let _rundate = %sysfunc(date(), yymmddn8.);

    /* Validate required parameters */
    %if %length(&newLib) = 0 or %length(&oldLib) = 0 or %length(&outLib) = 0 %then %do;
        %put ERROR: versionDiff — newLib, oldLib, and outLib are all required (pre-assigned librefs);
        %return;
    %end;
    %if %length(&keyvar) = 0 %then %do;
        %put WARNING: versionDiff — keyvar is recommended for correct merge matching;
    %end;

    %put NOTE: versionDiff — comparing %upcase(&newLib) (new) vs %upcase(&oldLib) (old);
    %put NOTE: versionDiff — output library: %upcase(&outLib);

    /* Get list of datasets from new library */
    proc sql noprint;
        select memname into :_dslist separated by ' '
        from dictionary.tables
        where upcase(libname) = "%upcase(&newLib)" and memtype = "DATA";
    quit;

    %if %length(&_dslist) = 0 %then %do;
        %put ERROR: versionDiff — no datasets found in %upcase(&newLib);
        %return;
    %end;

    %let _dsetnum = 1;

    /* ---- Loop over each dataset ---- */
    %do %while (%scan(&_dslist, &_dsetnum) ne );

        %let _dset = %scan(&_dslist, &_dsetnum);

        /* Step 1: Copy new dataset to outLib with dropVar removed */
        data &outLib..&_dset;
            set &newLib..&_dset;
            %if %length(&dropVar) > 0 %then %do;
                drop &dropVar;
            %end;
        run;

        /* Step 2: Load old dataset, filter out Inact */
        data &_dset;
            set &oldLib..&_dset;
            %if %length(&dropVar) > 0 %then %do;
                drop &dropVar;
            %end;
            if _STATUE = 'Inact' then delete;
        run;

        /* Step 3: Get variable lists for comparison */
        proc contents noprint data=&outLib..&_dset
            out=_Var(keep=MEMNAME VARNUM NAME where=(upcase(NAME) ^in (&exclVar)));
        run;
        proc contents noprint data=&_dset
            out=_OldVar(keep=MEMNAME VARNUM NAME where=(upcase(NAME) ^in (&exclVar)));
        run;

        proc sort data=_Var;    by MEMNAME NAME; run;
        proc sort data=_OldVar; by MEMNAME NAME; run;

        /* Step 4: Sort by key variables */
        %if %length(&keyvar) > 0 %then %do;
            proc sort data=&outLib..&_dset out=&outLib..&_dset; by &keyvar; run;
            proc sort data=&_dset      out=&_dset;             by &keyvar; run;
        %end;

        /* Step 5: Identify records only in old (Inact) */
        data T2_&_dset;
            %if %length(&keyvar) > 0 %then %do;
                merge &_dset(in=a)
                      &outLib..&_dset(in=b);
                by &keyvar;
            %end;
            %else %do;
                merge &_dset(in=a)
                      &outLib..&_dset(in=b);
            %end;
            if a and ^b;
            _Statue = "Inact";
            RUNDATE = "&_rundate.";
        run;

        /* Step 6: Build rename/DIFF variable lists from metadata */
        data _null_;
            merge _Var(in=a) _OldVar(in=b);
            by MEMNAME NAME;
            if a;
            length RenameVar DIFFVar NewVar $10000;
            retain RenameVar DIFFVar NewVar;
            call symput('_Col'||left(put(_n_,3.)), compress(VARNUM));
            call symput('_VarName'||left(put(_n_,3.)), compress(NAME));
            if first.MEMNAME then do;
                RenameVar = ""; DIFFVar = ""; NewVar = " ";
            end;
            if a and b then
                RenameVar = strip(strip(RenameVar)||" "||strip(NAME)||"=last_"||strip(NAME));
            if a and ^b then
                NewVar = strip(strip(NewVar)||"last_"||strip(NAME))||"=''; ";
            DIFFVar = strip(strip(DIFFVar)||" DIFF"||strip(VARNUM));
            if last.MEMNAME then do;
                call symput("_RenameVar", RenameVar);
                call symput("_NewVar",    NewVar);
                call symput("_DIFFVar",   DIFFVar);
                call symput('_number', left(compress(put(_n_,best.))));
            end;
        run;

        /* Step 7: Rename old-dataset variables and append */
        %if %length(&keyvar) > 0 %then %do;
            proc sort data=&_dset out=&oldLib._&_dset(rename=(&_RenameVar)); by &keyvar; run;
        %end;
        %else %do;
            proc sort data=&_dset out=&oldLib._&_dset(rename=(&_RenameVar)); run;
        %end;
        data &oldLib._&_dset;
            set &oldLib._&_dset;
            &_NewVar;
        run;

        %if %length(&keyvar) > 0 %then %do;
            proc sort data=&outLib..&_dset out=&newLib._&_dset; by &keyvar; run;
        %end;
        %else %do;
            proc sort data=&outLib..&_dset out=&newLib._&_dset; run;
        %end;

        /* Step 8: Detect changes — new, update, no change */
        data T_&_dset(keep=&keyvar Type)
             C_&_dset(keep=&keyvar &_DIFFVar);
            %if %length(&keyvar) > 0 %then %do;
                merge &oldLib._&_dset(in=a)
                      &newLib._&_dset(in=b);
                by &keyvar;
            %end;
            %else %do;
                merge &oldLib._&_dset(in=a)
                      &newLib._&_dset(in=b);
            %end;
            length Type &_DIFFVar $20;
            if b then do;
                %do i = 1 %to &_number;
                    DIFF&&_Col&i.. = "";
                    if (upcase(&&_VarName&i..) ^= upcase(last_&&_VarName&i..) & a) then do;
                        DIFF&&_Col&i.. = compress("C"||&&_Col&i..);
                        Type = "Update";
                    end;
                %end;
                if ^a then Type = "New";
                output;
            end;
        run;

        /* Step 9: Build final version dataset */
        data T1_&_dset(drop=last_RUNDATE Type);
            %if %length(&keyvar) > 0 %then %do;
                merge &newLib._&_dset(in=a)
                      &oldLib._&_dset(keep=&keyvar RUNDATE rename=(RUNDATE=last_RUNDATE))
                      T_&_dset;
                by &keyvar;
            %end;
            %else %do;
                merge &newLib._&_dset(in=a)
                      &oldLib._&_dset(keep=RUNDATE rename=(RUNDATE=last_RUNDATE))
                      T_&_dset;
            %end;
            if a;
            _Statue = Type;
            if Type ^in ("New","Update") then RUNDATE = last_RUNDATE;
            if RUNDATE = '' then RUNDATE = "&_rundate.";
        run;

        data &outLib..&_dset;
            set T1_&_dset T2_&_dset;
        run;

        /* Step 10: Frequency summary */
        proc sort data=&outLib..&_dset out=_freq; by &keyvar _Statue; run;

        PROC FREQ DATA=_freq noprint;
            TABLES _Statue / OUT=_freqOut;
        RUN;

        /* Count distinct subjects for Subject Total */
        proc sort data=&outLib..&_dset(where=(usubjid ^= '')) out=_subj nodupkey;
            by usubjid;
        run;

        data _subj1;
            set _subj;
            length A $20;
            A = "Subject Total";
        run;

        PROC FREQ DATA=_subj1 noprint;
            TABLES A / OUT=_freqSubj;
        RUN;

        data _freqSubj;
            set _freqSubj;
            rename A = _Statue;
            if A = '' then A = "Subject Total";
        run;

        data _freqAll;
            set _freqSubj _freqOut;
            length DatasetName $32;
            DatasetName = "&_dset";
        run;

        /* Ensure all status categories exist */
        %local _cats;
        %let _cats = Inact Update New Subject_Total stable;
        %do i = 1 %to 5;
            %let _cat = %scan(&_cats, &i);
            proc sql noprint;
                select count(*) into :_cnt from _freqAll
                %if &_cat = Subject_Total %then
                    where _Statue = 'Subject Total';
                %else %if &_cat = stable %then
                    where _Statue = '';
                %else
                    where _Statue = "&_cat";
                ;
            quit;
            %if &_cnt = 0 %then %do;
                proc sql noprint;
                    %if &_cat = stable %then
                        insert into _freqAll (_Statue, COUNT, PERCENT, DatasetName)
                        values ('total', 0, 0, "&_dset");
                    %else
                        insert into _freqAll (_Statue, COUNT, PERCENT, DatasetName)
                        values ("&_cat", 0, 0, "&_dset");
                    ;
                quit;
            %end;
        %end;

        data _freqAll;
            set _freqAll;
            if _Statue = '' then _Statue = 'stable';
        run;

        proc transpose data=_freqAll
            out=_freqTrans(drop=_NAME_)
            prefix=Form_;
            id _Statue;
            by DatasetName;
            var COUNT;
        run;

        data _freqTrans(drop=Form_stable);
            retain DatasetName _LABEL_ Form_New Form_Update Form_Inact Form_Total Form_Subject_Total;
            set _freqTrans;
            if Form_stable = '' then Form_stable = 0;
            Form_Total = Form_Inact + Form_Update + Form_New + Form_stable;
        run;

        proc append base=_SourceSummary data=_freqTrans(drop=_LABEL_); run;

        %let _dsetnum = %eval(&_dsetnum + 1);
    %end;

    /* ---- Export to Excel ---- */

    /* Build source description */
    %let _libpath = %sysfunc(pathname(&newLib));
    %let _filename = %sysfunc(scan(&_libpath, -2, '\'))\%sysfunc(scan(&_libpath, -1, '\'));

    data _src1;
        length sourcetype filename $256;
        sourcetype = "SASdataset New";
        filename   = "&_filename";
        output;
    run;

    %let _libpath = %sysfunc(pathname(&oldLib));
    %let _filename = %sysfunc(scan(&_libpath, -2, '\'))\%sysfunc(scan(&_libpath, -1, '\'));

    data _src2;
        length sourcetype filename $256;
        sourcetype = "SASdataset Old";
        filename   = "&_filename";
        output;
    run;

    data _sourceInfo;
        set _src1 _src2;
    run;

    /* Export to current working directory */
    %let _xlsxname = Compare&_rundate.;
    %put NOTE: versionDiff — writing &_xlsxname..xlsx;

    PROC EXPORT DATA=_SourceSummary
        OUTFILE= "&_xlsxname..xlsx"
        DBMS=xlsx REPLACE;
        SHEET="SourceSummary";
    RUN;

    PROC EXPORT DATA=_sourceInfo
        OUTFILE= "&_xlsxname..xlsx"
        DBMS=xlsx REPLACE;
        SHEET="sourcetype";
    RUN;

    /* Export each dataset from outLib */
    %let _dsetnum = 1;
    %do %while (%scan(&_dslist, &_dsetnum) ne );
        %let _dset = %scan(&_dslist, &_dsetnum);
        PROC EXPORT DATA=&outLib..&_dset
            OUTFILE= "&_xlsxname..xlsx"
            DBMS=xlsx REPLACE label;
            SHEET=&_dset;
        RUN;
        %let _dsetnum = %eval(&_dsetnum + 1);
    %end;

    /* Cleanup */
    %_dropIfExists(dsn=_SourceSummary);
    %_dropIfExists(dsn=_sourceInfo);
    %_dropIfExists(dsn=_src1);
    %_dropIfExists(dsn=_src2);
    %_dropIfExists(dsn=_freqOut);
    %_dropIfExists(dsn=_freqSubj);
    %_dropIfExists(dsn=_freqAll);
    %_dropIfExists(dsn=_freqTrans);
    %_dropIfExists(dsn=_freq);
    %_dropIfExists(dsn=_subj);
    %_dropIfExists(dsn=_subj1);
    %_dropIfExists(dsn=_Var);
    %_dropIfExists(dsn=_OldVar);

    %put NOTE: versionDiff — complete.  Output: &_xlsxname..xlsx;
%mend versionDiff;


/* ========== MAIN MACRO: versionColor ========== */

%macro versionColor(outLib=);

    %if %length(&outLib) = 0 %then %do;
        %put ERROR: versionColor — outLib (pre-assigned libref) is required;
        %return;
    %end;

    %put NOTE: versionColor — applying colour highlights via DDE;
    %put NOTE: versionColor — ensure the Excel workbook is OPEN before running;

    /* Get list of datasets in the output library */
    proc sql noprint;
        select memname into :_vclist separated by ' '
        from dictionary.tables
        where upcase(libname) = "%upcase(&outLib)" and memtype = "DATA";
    quit;

    %local _vcnum _vcset;
    %let _vcnum = 1;

    %do %while (%scan(&_vclist, &_vcnum) ne );
        %let _vcset = %scan(&_vclist, &_vcnum);

        /* Check if C_ dataset exists (contains diff info) */
        %if %sysfunc(exist(C_&_vcset, data)) %then %do;

            /* Activate the sheet in Excel via DDE */
            filename cmds dde 'excel|system';
            data _null_;
                file cmds;
                put '[workbook.select("' "&_vcset" '")]';
            run;

            /* Highlight changed cells */
            filename sas2xl dde 'excel|system';
            data _null_;
                file sas2xl;
                set C_&_vcset;
                array _chr_ _CHAR_;
                do over _chr_;
                    if index(upcase(vname(_chr_)), "DIFF") and ^missing(_chr_) then do;
                        _range = compress('r'||_n_+1||_chr_||':r'||_n_+1||_chr_);
                        put '[select("' _range '")]';
                        put '[patterns(1,,37)]';
                    end;
                end;
            run;

        %end;

        %let _vcnum = %eval(&_vcnum + 1);
    %end;

    %put NOTE: versionColor — complete.  Changed cells highlighted in Excel.;
%mend versionColor;
