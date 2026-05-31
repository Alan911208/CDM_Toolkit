/****************************************************************************
 *  Macro: %dmQueryReport
 *  Category: Reporting
 *  Purpose: Generate the DMR Query Summary tables (T5.1-T5.3) in RTF format
 *           from query detail and form metadata Excel sheets.
 *
 *  T5.1 — Query counts by form, top-4 query texts
 *  T5.2 — Response-time statistics by query type
 *  T5.3 — Site-level key findings
 *
 *  Supports two template versions:
 *    ver=2 — legacy column mapping (older DMR exports)
 *    ver=3 — v3.0 column mapping (current standard)
 *
 *  Note: This macro reads Excel sheets directly via PROC IMPORT,
 *        so no dependency on %xls2sas is needed.
 ****************************************************************************/


/* ========== INTERNAL HELPERS ========== */

%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;

%macro _dmImportExcel(file=, sheet=, out=);
    %if %sysfunc(exist("&file")) %then %do;
        proc import datafile="&file"
            out=&out dbms=xlsx replace;
            sheet="&sheet"; getnames=no;
        run;
    %end;
    %else %do;
        %put ERROR: _dmImportExcel — file &file not found.;
    %end;
%mend _dmImportExcel;


/* ========== MAIN MACRO ========== */

%macro dmQueryReport(
    qSheet  = ,    /* Excel sheet name containing query detail data */
    form    = ,    /* Excel sheet name containing form metadata */
    outRTF  = ,    /* Output RTF file path */
    ver     = 3    /* Template version: 2 (legacy) or 3 (v3.0) */
);

    %local _nTotal _nSite _nSubj;

    /* ---- Validate required parameters ---- */
    %if %length(&qSheet) = 0 or %length(&form) = 0 %then %do;
        %put ERROR: dmQueryReport — qSheet and form are required parameters;
        %return;
    %end;
    %if %length(&outRTF) = 0 %then %do;
        %put ERROR: dmQueryReport — outRTF is required;
        %return;
    %end;
    %if &ver not in (2, 3) %then %do;
        %put ERROR: dmQueryReport — ver must be 2 or 3 (got &ver);
        %return;
    %end;

    %put NOTE: dmQueryReport — importing query data (ver=&ver);
    %put NOTE: dmQueryReport — output: &outRTF;

    /* ============================================ */
    /*  Import query and form metadata              */
    /* ============================================ */

    %_dmImportExcel(file=&qSheet, sheet=&qSheet, out=_query);
    %_dmImportExcel(file=&form,   sheet=&form,   out=_qForm);

    /* ---- Verify import succeeded ---- */
    %if %sysfunc(exist(_query, data)) = 0 or %sysfunc(exist(_qForm, data)) = 0 %then %do;
        %put ERROR: dmQueryReport — import failed; check that qSheet and form point to valid Excel data;
        %return;
    %end;

    /* ============================================ */
    /*  Template-specific column mapping            */
    /* ============================================ */

    %if &ver = 2 %then %do;
        /* ---- Ver 2: legacy format ---- */
        data _q2;
            set _qForm;
            keep _c1_ _c2_;
            rename _c1_ = c
                   _c2_ = _c8_;
        run;

        proc sort data=_q2;    by _c8_; run;
        proc sort data=_query; by _c8_; run;

        data _query;
            merge _query _q2;
            by _c8_;
            if _c1_ = '' then delete;
        run;

        data _query;
            set _query(where=(_c21_ not in ("Cancel","Cancelled")));
            rename _c3_  = site_name
                   _c4_  = subject_No
                   _c8_  = form_name
                   c     = Field_OID
                   _c13_ = query_name
                   _c14_ = Query_Open_Date
                   _c17_ = query_text
                   _c18_ = Answer_Date
                   _c21_ = Query_Status
                   _c22_ = Resolve_Date;
        run;
    %end;
    %else %if &ver = 3 %then %do;
        /* ---- Ver 3: current standard format ---- */
        data _q2;
            set _qForm;
            keep _c1_ _c2_;
            rename _c1_ = c
                   _c2_ = _c9_;
        run;

        proc sort data=_q2;    by _c9_; run;
        proc sort data=_query; by _c9_; run;

        data _query;
            merge _query _q2;
            by _c9_;
            if _c2_ = '' then delete;
        run;

        data _query;
            set _query(where=(_c22_ not in ("Cancel","Cancelled")));
            rename _c4_  = site_name
                   _c5_  = subject_No
                   _c9_  = form_name
                   c     = Field_OID
                   _c14_ = query_name
                   _c15_ = Query_Open_Date
                   _c18_ = query_text
                   _c19_ = Answer_Date
                   _c22_ = Query_Status
                   _c23_ = Resolve_Date;
        run;
    %end;

    /* ============================================ */
    /*  Summary statistics                          */
    /* ============================================ */

    proc sql noprint;
        select count(*),
               count(distinct site_name),
               count(distinct subject_No)
        into :_nTotal, :_nSite, :_nSubj
        from _query
        where Query_Status not in ("Cancel","Cancelled");
    quit;

    %put NOTE: dmQueryReport — Total queries: &_nTotal;
    %put NOTE: dmQueryReport — Sites: &_nSite, Subjects: &_nSubj;

    /* ---- T5.1: Query counts by form with top-4 query texts ---- */
    /* Count queries per form */
    proc sql;
        create table _t5_1_base as
            select form_name,
                   count(*) as n_queries format=comma12.,
                   count(distinct site_name)  as n_sites,
                   count(distinct subject_No) as n_subjects
            from _query
            where Query_Status not in ("Cancel","Cancelled")
            group by form_name
            order by n_queries descending;
    quit;

    /* Get top 4 query texts per form */
    proc sql;
        create table _t5_1_texts as
            select form_name, query_text, count(*) as n
            from _query
            where Query_Status not in ("Cancel","Cancelled")
            group by form_name, query_text
            order by form_name, n descending;
    quit;

    /* Select top 4 per form */
    data _t5_1_top4;
        set _t5_1_texts;
        by form_name;
        retain _rank;
        if first.form_name then _rank = 0;
        _rank + 1;
        if _rank <= 4;
        keep form_name query_text n _rank;
        label _rank = 'Rank';
    run;

    /* ---- T5.2: Response-time statistics by query type ---- */
    /* Calculate response time = Answer_Date - Query_Open_Date */
    data _t5_2;
        set _query;
        where Query_Status not in ("Cancel","Cancelled")
              and Answer_Date ^= '' and Query_Open_Date ^= '';
        _open   = input(compress(Query_Open_Date, '-/ '), yymmdd10.);
        _answer = input(compress(Answer_Date, '-/ '), yymmdd10.);
        if ^missing(_open) and ^missing(_answer) then do;
            _days = _answer - _open;
            if _days >= 0;
            output;
        end;
        drop _open _answer;
    run;

    proc sql;
        create table _t5_2_stats as
            select query_name,
                   count(*)          as n,
                   mean(_days)       as mean_days  format=6.1,
                   median(_days)     as median_days format=6.0,
                   min(_days)        as min_days   format=6.0,
                   max(_days)        as max_days   format=6.0
            from _t5_2
            group by query_name
            order by n descending;
    quit;

    /* ---- T5.3: Site-level summary ---- */
    proc sql;
        create table _t5_3 as
            select site_name,
                   count(*)                             as n_queries format=comma12.,
                   count(distinct form_name)            as n_forms,
                   count(distinct subject_No)           as n_subjects,
                   count(distinct Query_Open_Date)      as n_open_dates,
                   sum(case when Query_Status = 'Answered' then 1 else 0 end)
                       as n_answered,
                   sum(case when Query_Status = 'Open' then 1 else 0 end)
                       as n_open
            from _query
            where Query_Status not in ("Cancel","Cancelled")
            group by site_name
            order by n_queries descending;
    quit;

    /* ============================================ */
    /*  Output to RTF                               */
    /* ============================================ */

    ods rtf file="&outRTF" style=journal;

    title1 "DMR Query Summary Report";
    title2 "Generated: %sysfunc(date(), worddate.)";
    footnote1 "Total: &_nTotal queries, &_nSubj subjects, &_nSite sites";

    /* ---- T5.1: Form-Level Query Counts ---- */
    title3 "Table 5.1: Query Counts by Form";
    proc report data=_t5_1_base nowd;
        column form_name n_queries n_sites n_subjects;
        define form_name / "Form" width=30;
        define n_queries / "Number of Queries" format=comma12.;
        define n_sites   / "Distinct Sites"    format=comma12.;
        define n_subjects/ "Distinct Subjects" format=comma12.;
    run;

    title4 "Top 4 Query Texts per Form";
    proc report data=_t5_1_top4 nowd;
        column form_name _rank query_text n;
        define form_name  / "Form";
        define _rank      / "Rank";
        define query_text / "Query Text" width=60;
        define n          / "Count" format=comma12.;
        break after form_name / skip;
    run;

    /* ---- T5.2: Response Time ---- */
    title3 "Table 5.2: Response Time by Query Type";
    proc report data=_t5_2_stats nowd;
        column query_name n mean_days median_days min_days max_days;
        define query_name  / "Query Type";
        define n           / "N";
        define mean_days   / "Mean (days)";
        define median_days / "Median (days)";
        define min_days    / "Min";
        define max_days    / "Max";
    run;

    /* ---- T5.3: Site Key Findings ---- */
    title3 "Table 5.3: Site-Level Key Findings";
    proc report data=_t5_3 nowd;
        column site_name n_queries n_forms n_subjects n_answered n_open;
        define site_name   / "Site";
        define n_queries   / "Total Queries" format=comma12.;
        define n_forms     / "Forms with Queries";
        define n_subjects  / "Subjects with Queries";
        define n_answered  / "Answered";
        define n_open      / "Open";
    run;

    ods rtf close;

    %put NOTE: dmQueryReport — RTF output written to &outRTF;

    /* ---- Cleanup ---- */
    %_dropIfExists(dsn=_query);
    %_dropIfExists(dsn=_qForm);
    %_dropIfExists(dsn=_q2);
    %_dropIfExists(dsn=_t5_1_base);
    %_dropIfExists(dsn=_t5_1_texts);
    %_dropIfExists(dsn=_t5_1_top4);
    %_dropIfExists(dsn=_t5_2);
    %_dropIfExists(dsn=_t5_2_stats);
    %_dropIfExists(dsn=_t5_3);

    %put NOTE: dmQueryReport — complete.;
%mend dmQueryReport;
