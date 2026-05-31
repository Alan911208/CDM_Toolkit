/**********************************************************************************
*  CDM Toolkit — rawdata_export.sas
*  Description : Review raw SAS datasets, optimize char lengths, generate Excel report
*  Usage       : Upload .sas7bdat files via Web UI, libname raw is auto-configured
*  Output      : rawdata_content.xlsx (in _cdm_tmpdir or current path)
*  SAS config  : zh (GBK) or u8 (UTF-8) — both supported
**********************************************************************************/

dm "log;clear;";

/* ── Determine output path ── */
%let _outdir = %sysfunc(pathname(work));
%if %symexist(_cdm_tmpdir) %then %do;
  %let _outdir = &_cdm_tmpdir;
%end;

/* ── Verify raw libname exists ── */
%let _raw_exists = 0;
data _null_;
  if libref("RAW") then call symputx("_raw_exists", 0);
  else call symputx("_raw_exists", 1);
run;
%if &_raw_exists = 0 %then %do;
  %put ERROR: libname RAW is not assigned. Please upload SAS datasets first.;
  data _null_; put "ERROR: libname RAW not assigned"; run;
  %goto exit;
%end;

/* ── Get dataset list from raw library ── */
proc contents data=raw._all_ out=_content_raw noprint;
run;

proc sql noprint;
  select distinct memname into :ds_list separated by ' '
  from _content_raw;
  select count(distinct memname) into :nds
  from _content_raw;
quit;

%if &nds = 0 %then %do;
  %put WARNING: No SAS datasets found in RAW library.;
  data _null_; put "WARNING: No datasets found"; run;
  %goto exit;
%end;
%put INFO: Found &nds dataset(s): &ds_list;

/* ── Copy datasets from raw to work ── */
proc datasets library=raw nolist;
  copy in=raw out=work;
  select &ds_list;
quit;

/* ── Optimize character variable lengths ── */
%macro _optimize_char_lengths(lib=work);
  %local i ds n_char;
  %let i = 1;
  %let ds = %scan(&ds_list, &i, %str( ));

  %do %while (&ds ne );
    proc sql noprint;
      select count(*) into :_nobs from &lib..&ds;
    quit;

    %if &_nobs > 0 %then %do;
      /* Get dataset label */
      data _null_;
        set sashelp.vtable(where=(libname="%upcase(&lib)" and memname="%upcase(&ds)")) end=eof;
        call symputx("_dslabel", memlabel);
        if eof and missing(memlabel) then call symputx("_dslabel", "");
      run;

      /* Get char variable names */
      proc contents data=&lib..&ds out=_ctmp varnum noprint;
      run;

      proc sql noprint;
        select count(*) into :n_char from _ctmp where type=2;
        select name into :_cv1-:_cv%left(&n_char) from _ctmp where type=2;
        select name into :_allvars separated by ' ' from _ctmp order by varnum;
      quit;

      %if &n_char > 0 %then %do;
        /* Compute max lengths */
        proc sql;
          create table _maxlen as
          select
            %do j = 1 %to &n_char;
              max(length(&&_cv&j)) as &&_cv&j
              %if &j < &n_char %then ,;
            %end;
          from &lib..&ds;
        quit;

        proc transpose data=_maxlen out=_maxlen_t;
          var _all_;
        run;

        data _null_;
          set _maxlen_t;
          call symputx(cats("_newlen", put(_N_, best.)),
            cats(_name_, " $", strip(put(col1, best.))));
        run;

        options varlenchk=nowarn;
        data &lib..&ds %if %length(&_dslabel)>0 %then (label="&_dslabel");;
          retain &_allvars;
          length
            %do j = 1 %to &n_char;
              &&_newlen&j
            %end;
          ;
          set &lib..&ds;
          informat _character_;
          format _character_;
        run;
        proc datasets nolist;
          delete _ctmp _maxlen _maxlen_t;
        quit;
      %end;
    %end;

    %let i = %eval(&i + 1);
    %let ds = %scan(&ds_list, &i, %str( ));
  %end;
%mend _optimize_char_lengths;

%_optimize_char_lengths;

/* ── Build content metadata ── */
proc contents data=work._all_ out=_content2 noprint;
run;

proc sort data=_content2;
  by memname varnum;
run;

data _content_final;
  set _content2;
  where memname in (%do i=1 %to &nds; "%scan(&ds_list, &i, %str( ))" %if &i < &nds %then ,; %end;);

  /* Filter out internal/system variables */
  if prxmatch("/^__|_U$/", strip(name)) then delete;

  length variable_name $32 variable_label $256 vartype $4;
  variable_name = name;

  if prxmatch("/Specify/i", label) then
    variable_label = catx(" ", strip(scan(label, 1, '(')), "(Specify)");
  else
    variable_label = strip(scan(label, 1, '('));

  if type = 2 then vartype = "Char";
  else vartype = "Num";

  keep memname variable_name variable_label vartype length varnum;
run;

proc sort data=_content_final;
  by memname varnum;
run;

/* ── Extract dataset-level FORM info ── */
data _toc;
  length rawdata $32 form $200;
  stop;
run;

%macro _get_form;
  %local k ds;
  %let k = 1;
  %let ds = %scan(&ds_list, &k, %str( ));
  %do %while (&ds ne );
    %let _formval = ;
    data _null_;
      if 0 then set work.&ds nobs=_n;
      call symputx("_hasobs", _n > 0);
      dsid = open("work.&ds");
      if dsid then do;
        _fnum = varnum(dsid, "FORM");
        if _fnum > 0 then do;
          rc = fetchobs(dsid, 1);
          if rc = 0 then call symputx("_formval", getvarc(dsid, _fnum));
        end;
        rc = close(dsid);
      end;
    run;

    data _toc;
      set _toc;
      rawdata = "&ds";
      form = symget("_formval");
      output;
    run;

    %let k = %eval(&k + 1);
    %let ds = %scan(&ds_list, &k, %str( ));
  %end;
%mend _get_form;

%_get_form;

/* ── Generate Excel report ── */
options papersize=letter orientation=landscape nodate nonumber center missing=" " nobyline;
ods escapechar="@";
ods html close;
ods excel file="&_outdir.\rawdata_content.xlsx" style=normal;

/* Sheet 1: Dataset overview */
ods excel options(sheet_name="CONTENT" absolute_row_height='16'
  absolute_column_width='10,40,10' autofilter='column');

proc report data=_toc missing center nowd headline headskip split='~'
  style(header)={just=left asis=on nobreakspace=on color=black
    backgroundcolor=#DCE6F1 font=("Arial", 10pt) font_weight=bold};
  column rawdata form;
  define rawdata / display 'Raw Dataset';
  define form    / display 'Form';
run;

/* Sheet per dataset: variable listing */
%macro _report_sheets;
  %local k ds;
  %let k = 1;
  %let ds = %scan(&ds_list, &k, %str( ));
  %do %while (&ds ne );
    ods excel options(sheet_name="%upcase(&ds)"
      absolute_row_height='16'
      absolute_column_width='15,50,8,8,20,10,5,20,10,12'
      autofilter='column');

    proc report data=_content_final(where=(memname="%upcase(&ds)"))
      missing center nowd headline headskip split='~'
      style(header)={just=left asis=on nobreakspace=on color=black
        backgroundcolor=#DCE6F1 font=("Arial", 10pt) font_weight=bold};

      column variable_name variable_label vartype length varnum;
      define variable_name  / display 'Variable_Name';
      define variable_label / display 'Variable_Label';
      define vartype        / display 'Type';
      define length         / display 'Length';
      define varnum         / display 'Variable_Order';
    run;

    %let k = %eval(&k + 1);
    %let ds = %scan(&ds_list, &k, %str( ));
  %end;
%mend _report_sheets;

%_report_sheets;

ods excel close;
ods html;

/* ── Cleanup ── */
proc datasets library=work nolist;
  delete _content_raw _content2 _content_final _toc;
quit;

%put INFO: rawdata_content.xlsx generated at &_outdir;
%exit:
