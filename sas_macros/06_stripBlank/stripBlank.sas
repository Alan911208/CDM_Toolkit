/*****************************************************************************
 * Tool 6:  %stripBlank
 * Purpose: Remove blank/duplicate placeholder rows from CDM page-based
 *          datasets. Handles pStatus, seqCD/seqTEXT cleanup, and auto-fills
 *          missing index values.
 *
 * Parameters:
 *   inlib   = Input library (default: work)
 *   outlib  = Output library (default: same as inlib)
 *   insets  = Input dataset name(s), space-separated (REQUIRED)
 *   outsets = Output dataset name(s), space-separated (default: same as insets)
 *   idxVar  = Index variable name (default: SEQNUM)
 *   keepPS  = Keep pStatus column? (NO by default)
 *
 * Usage:
 *   %stripBlank(inlib=raw, outlib=clean, insets=lb1 lb2, outsets=lb1_clean lb2_clean);
 *   %stripBlank(insets=ae1, idxVar=LBTESTCD);
 *   %stripBlank(inlib=work, outlib=derived, insets=dm, keepPS=YES);
 *
 * Source: DM_Toolbox.sas lines 452-536
 *****************************************************************************/

%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;

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
