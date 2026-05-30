/****************************************************************************
 *  Test Script: %dmQueryReport
 *  Purpose: Create mock query and form metadata datasets, export them as
 *           temporary Excel files, and run dmQueryReport with ver=3.
 *
 *  Mock data:
 *    - 3 sites (Site A, Site B, Site C)
 *    - 3 forms (DM, AE, CM)
 *    - Multiple queries with various statuses, open dates, and answer dates
 *    - Some Cancel/Cancelled queries (should be excluded)
 *
 *  Note: This test creates temporary Excel files via PROC EXPORT since
 *        the macro uses PROC IMPORT to read Excel.
 ****************************************************************************/

/* ---- Clean workspace ---- */
proc datasets library=work nolist kill; quit;

/* ---- Build mock QUERY dataset (ver=3 format) ---- */
/* Columns: _c1_ _c2_ _c3_ _c4_ _c5_ _c6_ _c7_ _c8_ _c9_ _c10_
   _c11_ _c12_ _c13_ _c14_ _c15_ _c16_ _c17_ _c18_ _c19_ _c20_
   _c21_ _c22_ _c23_
   Key: _c4_=site, _c5_=subject, _c9_=form_key, _c14_=query_name,
        _c15_=OpenDate, _c18_=query_text, _c19_=AnswerDate,
        _c22_=status, _c23_=ResolveDate
*/
data _query_raw;
    length _c1_-_c23_ $50;
    array c_{23} $50 _c1_-_c23_;
    /* Initialize all to empty */
    do i = 1 to 23; c_{i} = ''; end;

    /* Row 1: DM form, Normal query, Answered */
    _c4_ = 'Site A'; _c5_ = '001'; _c9_ = 'DM_FORM';
    _c14_ = 'Auto Query'; _c15_ = '2024-01-15'; _c18_ = 'Missing value for BIRTHDT';
    _c19_ = '2024-01-17'; _c22_ = 'Answered'; _c23_ = '2024-01-17'; output;

    /* Row 2: AE form, Normal query, Answered */
    _c4_ = 'Site A'; _c5_ = '001'; _c9_ = 'AE_FORM';
    _c14_ = 'Manual Query'; _c15_ = '2024-01-20'; _c18_ = 'AE term unclear';
    _c19_ = '2024-01-25'; _c22_ = 'Answered'; _c23_ = '2024-01-25'; output;

    /* Row 3: CM form, Cancelled — should be excluded */
    _c4_ = 'Site A'; _c5_ = '002'; _c9_ = 'CM_FORM';
    _c14_ = 'Auto Query'; _c15_ = '2024-02-01'; _c18_ = 'Duplicate entry';
    _c19_ = ''; _c22_ = 'Cancelled'; _c23_ = ''; output;

    /* Row 4: AE form, Open query */
    _c4_ = 'Site B'; _c5_ = '003'; _c9_ = 'AE_FORM';
    _c14_ = 'Manual Query'; _c15_ = '2024-02-10'; _c18_ = 'Verify AE severity';
    _c19_ = ''; _c22_ = 'Open'; _c23_ = ''; output;

    /* Row 5: DM form, Cancel — should be excluded */
    _c4_ = 'Site B'; _c5_ = '004'; _c9_ = 'DM_FORM';
    _c14_ = 'Auto Query'; _c15_ = '2024-02-15'; _c18_ = 'Already resolved';
    _c19_ = ''; _c22_ = 'Cancel'; _c23_ = ''; output;

    /* Row 6: CM form, Answered */
    _c4_ = 'Site B'; _c5_ = '005'; _c9_ = 'CM_FORM';
    _c14_ = 'Auto Query'; _c15_ = '2024-03-01'; _c18_ = 'CM dose missing';
    _c19_ = '2024-03-03'; _c22_ = 'Answered'; _c23_ = '2024-03-03'; output;

    /* Row 7: DM form, Answered (same form as row 1, different subject) */
    _c4_ = 'Site C'; _c5_ = '006'; _c9_ = 'DM_FORM';
    _c14_ = 'Manual Query'; _c15_ = '2024-03-05'; _c18_ = 'Missing value for BIRTHDT';
    _c19_ = '2024-03-10'; _c22_ = 'Answered'; _c23_ = '2024-03-10'; output;

    /* Row 8: DM form, different text */
    _c4_ = 'Site C'; _c5_ = '007'; _c9_ = 'DM_FORM';
    _c14_ = 'Auto Query'; _c15_ = '2024-03-10'; _c18_ = 'Date inconsistency: Visit vs Assessment';
    _c19_ = '2024-03-12'; _c22_ = 'Answered'; _c23_ = '2024-03-12'; output;

    /* Row 9: AE form, Open */
    _c4_ = 'Site C'; _c5_ = '007'; _c9_ = 'AE_FORM';
    _c14_ = 'Auto Query'; _c15_ = '2024-03-15'; _c18_ = 'AE term unclear';
    _c19_ = ''; _c22_ = 'Open'; _c23_ = ''; output;
run;

/* ---- Build mock FORM metadata dataset (ver=3 format) ---- */
/* _c1_ = Field_OID, _c2_ = form_key (_c9_ in query) */
data _qForm_raw;
    length _c1_ $50 _c2_ $50;
    _c1_ = 'BIRTHDT';  _c2_ = 'DM_FORM'; output;
    _c1_ = 'AETERM';   _c2_ = 'AE_FORM'; output;
    _c1_ = 'CMTRT';    _c2_ = 'CM_FORM'; output;
    _c1_ = 'AESEV';    _c2_ = 'AE_FORM'; output;
    _c1_ = 'AEOUT';    _c2_ = 'AE_FORM'; output;
    _c1_ = 'CMDOSFRQ'; _c2_ = 'CM_FORM'; output;
run;

/* ---- Export to temporary Excel files ---- */
%let _tmpDir = %sysfunc(pathname(work));

proc export data=_query_raw
    outfile="&_tmpDir\_test_query.xlsx"
    dbms=xlsx replace;
    sheet="_test_query";
run;

proc export data=_qForm_raw
    outfile="&_tmpDir\_test_form.xlsx"
    dbms=xlsx replace;
    sheet="_test_form";
run;

/* ---- Run %dmQueryReport ---- */
title "Test: %dmQueryReport — v3.0 DMR Query Summary";

%dmQueryReport(
    qSheet = &_tmpDir\_test_query.xlsx,
    form   = &_tmpDir\_test_form.xlsx,
    outRTF = &_tmpDir\DMR_Query_Test.rtf,
    ver    = 3
);

/* ---- Verify intermediate datasets ---- */

title2 "T5.1: Query counts by form";
%if %sysfunc(exist(_t5_1_base, data)) %then %do;
    proc print data=_t5_1_base noobs; run;
%end;

title2 "T5.1: Top 4 query texts per form";
%if %sysfunc(exist(_t5_1_top4, data)) %then %do;
    proc print data=_t5_1_top4 noobs; run;
%end;

title2 "T5.2: Response time by query type";
%if %sysfunc(exist(_t5_2_stats, data)) %then %do;
    proc print data=_t5_2_stats noobs; run;
%end;

title2 "T5.3: Site-level summary";
%if %sysfunc(exist(_t5_3, data)) %then %do;
    proc print data=_t5_3 noobs; run;
%end;

title "Test complete.";
