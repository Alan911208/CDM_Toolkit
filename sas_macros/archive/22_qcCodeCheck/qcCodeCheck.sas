/****************************************************************************
 *  Macro: %qcCodeCheck
 *  Category: Quality Control
 *  Purpose: Cross-validate MedDRA (AE) and WHODD (CM) coding data with
 *           7 automated consistency checks. Flagged records are written to
 *           an Excel workbook (Coding_QC.xls), one sheet per check.
 *
 *  AE Checks (MedDRA):
 *   AE-1: Same reported AE term   → different LLT codes
 *   AE-2: Same reported AE term   → different PT codes
 *   AE-3: Same LLT name           → different LLT codes
 *   AE-4: LLT name equals PT name → but codes differ
 *
 *  CM Checks (WHODD):
 *   CM-1: Same drug name         → different drug codes
 *   CM-2: Same drug name         → different preferred names
 *   CM-3: Same code + indication → different ATC4 selection
 *
 *  Note: This macro imports MedDRA and WHODD lookup sheets from Excel files
 *        via %xls2sas, or if %xls2sas is not present, falls back to PROC IMPORT.
 ****************************************************************************/


/* ========== INTERNAL HELPER ========== */

%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;

%macro _qcImportExcel(file=, sheet=, out=);
    %if %sysfunc(exist("&file")) %then %do;
        proc import datafile="&file"
            out=&out dbms=xlsx replace;
            sheet="&sheet"; getnames=no;
        run;
    %end;
    %else %do;
        %put WARNING: _qcImportExcel — file &file not found, skipping import;
    %end;
%mend _qcImportExcel;


/* ========== MAIN MACRO ========== */

%macro qcCodeCheck(
    project  = ,          /* Project ID to filter on           */
    out      = ,          /* Output directory for QC Excel     */
    aeFile   = ,          /* Path to MedDRA coding Excel file  */
    aeSheet  = MedDRA-DM-13-01,  /* AE sheet name */
    cmFile   = ,          /* Path to WHODD coding Excel file   */
    cmSheet  = WHODD-DM-13-02,   /* CM sheet name */
    aeTerm   = _c7_,      /* AE reported term column           */
    llt      = _c8_,      /* LLT name column                  */
    lltCode  = _c9_,      /* LLT code column                  */
    pt       = _c10_,     /* PT name column                   */
    ptCode   = _c11_,     /* PT code column                   */
    drug     = _c8_,      /* Drug name column (CM)            */
    ind      = _c9_,      /* Indication column (CM)           */
    drugCode = _c17_,     /* Drug code column (CM)            */
    atcName  = _c18_,     /* ATC name column (CM)             */
    atc4     = _c26_      /* ATC4 code column (CM)            */
);

    /* ---- Validate required parameters ---- */
    %if %length(&project) = 0 or %length(&out) = 0 %then %do;
        %put ERROR: qcCodeCheck — project and out are required parameters;
        %return;
    %end;

    %if %length(&aeFile) = 0 and %length(&cmFile) = 0 %then %do;
        %put WARNING: qcCodeCheck — no input files specified (aeFile, cmFile);
        %put WARNING: qcCodeCheck — macro will try to use pre-loaded _ae and _cm datasets;
    %end;

    %put NOTE: qcCodeCheck — project=&project, output=&out\Coding_QC.xls;

    /* ---- Step 0: Import AE and CM data ---- */
    %if %length(&aeFile) > 0 %then %do;
        %_qcImportExcel(file=&aeFile, sheet=&aeSheet, out=_ae);
    %end;
    %if %length(&cmFile) > 0 %then %do;
        %_qcImportExcel(file=&cmFile, sheet=&cmSheet, out=_cm);
    %end;

    /* ---- Verify that input datasets exist ---- */
    %if %sysfunc(exist(_ae, data)) = 0 %then %do;
        %put WARNING: qcCodeCheck — _ae dataset not found, skipping AE checks;
        %goto SKIP_AE;
    %end;
    %if %sysfunc(exist(_cm, data)) = 0 %then %do;
        %put WARNING: qcCodeCheck — _cm dataset not found, skipping CM checks;
        %goto SKIP_CM;
    %end;

    /* ============================================ */
    /*  PART A: MedDRA AE Coding Checks             */
    /* ============================================ */

    /* ---- Clean AE data: filter by project ---- */
    data _ae1;
        set _ae;
        if _c1_ ^= "&project" then delete;
        &aeTerm = compress(&aeTerm, '');
    run;

    /* AE-1: Same reported AE term, different LLT code */
    proc sort data=_ae1 out=_a1; by &aeTerm &lltCode; run;
    data _a2;
        set _a1;
        _bg = lag(&aeTerm);
        _prevLLTC = lag(&lltCode);
        length _fin $200;
        if &aeTerm = _bg and &lltCode ^= _prevLLTC then
            _fin = 'Same AE term, different LLT code — verify';
        if _fin ^= '' then output;
        drop _bg _prevLLTC;
    run;

    /* AE-2: Same reported AE term, different PT code */
    proc sort data=_ae1 out=_b1; by &aeTerm &ptCode; run;
    data _b2;
        set _b1;
        _bg = lag(&aeTerm);
        _prevPTC = lag(&ptCode);
        length _fin $200;
        if &aeTerm = _bg and &ptCode ^= _prevPTC then
            _fin = 'Same AE term, different PT code — verify';
        if _fin ^= '' then output;
        drop _bg _prevPTC;
    run;

    /* AE-3: Same LLT name, different LLT code */
    proc sort data=_ae1 out=_c1; by &llt &lltCode; run;
    data _c2;
        set _c1;
        _prevL  = lag(&llt);
        _prevLC = lag(&lltCode);
        length _fin $200;
        if &llt = _prevL and &lltCode ^= _prevLC then
            _fin = 'Same LLT, different LLT code — verify';
        if _fin ^= '' then output;
        drop _prevL _prevLC;
    run;

    /* AE-4: LLT name equals PT name, but codes differ */
    proc sort data=_ae1 out=_d1; by &llt &pt; run;
    data _d2;
        set _d1;
        length _fin $200;
        if &llt = &pt and &lltCode ^= &ptCode then
            _fin = 'LLT equals PT but codes differ — verify';
        if _fin ^= '' then output;
    run;

    %SKIP_AE:

    /* ============================================ */
    /*  PART B: WHODD CM Coding Checks              */
    /* ============================================ */

    %if %sysfunc(exist(_cm, data)) = 0 %then %goto SKIP_CM;

    /* ---- Clean CM data: filter by project ---- */
    data _cm1;
        set _cm;
        if _c1_ ^= "&project" then delete;
        &drug = compress(&drug, '');
    run;

    /* CM-1: Same drug name, different drug code */
    proc sort data=_cm1 out=_e1; by &drug &drugCode; run;
    data _e2;
        set _e1;
        _bg = lag(&drug);
        _prevDC = lag(&drugCode);
        length _fin $200;
        if &drug = _bg and &drugCode ^= _prevDC then
            _fin = 'Same drug name, different drug code — verify';
        if _fin ^= '' then output;
        drop _bg _prevDC;
    run;

    /* CM-2: Same drug name, different preferred name (ATC name) */
    proc sort data=_cm1 out=_f1; by &drug &atcName; run;
    data _f2;
        set _f1;
        _bg = lag(&drug);
        _prevPN = lag(&atcName);
        length _fin $200;
        if &drug = _bg and &atcName ^= _prevPN then
            _fin = 'Same drug name, different preferred name — verify';
        if _fin ^= '' then output;
        drop _bg _prevPN;
    run;

    /* CM-3: Same drug code + indication, different ATC4 */
    proc sort data=_cm1 out=_g1; by &drugCode &ind &atc4; run;
    data _g2;
        set _g1;
        _dc  = lag(&drugCode);
        _ind = lag(&ind);
        _a4  = lag(&atc4);
        length _fin $200;
        if &drugCode = _dc and &ind = _ind and &atc4 ^= _a4 then
            _fin = 'Same drug code + indication, different ATC — verify';
        if _fin ^= '' then output;
        drop _dc _ind _a4;
    run;

    %SKIP_CM:

    /* ============================================ */
    /*  Export to Excel workbook                    */
    /* ============================================ */

    %if %sysfunc(exist(_a2, data)) %then %do;
        libname _xl EXCEL "&out\Coding_QC.xls";
        data _xl."AE1_SameTerm_DiffLLTCode"n(dblabel=YES);   set _a2; run;
    %end;
    %else %do;
        libname _xl EXCEL "&out\Coding_QC.xls";
    %end;

    %if %sysfunc(exist(_b2, data)) %then %do;
        data _xl."AE2_SameTerm_DiffPTCode"n(dblabel=YES);   set _b2; run;
    %end;
    %if %sysfunc(exist(_c2, data)) %then %do;
        data _xl."AE3_SameLLT_DiffLLTCode"n(dblabel=YES);   set _c2; run;
    %end;
    %if %sysfunc(exist(_d2, data)) %then %do;
        data _xl."AE4_LLT_eq_PT_DiffCode"n(dblabel=YES);    set _d2; run;
    %end;
    %if %sysfunc(exist(_e2, data)) %then %do;
        data _xl."CM1_SameDrug_DiffDrugCode"n(dblabel=YES);  set _e2; run;
    %end;
    %if %sysfunc(exist(_f2, data)) %then %do;
        data _xl."CM2_SameDrug_DiffPrefName"n(dblabel=YES);  set _f2; run;
    %end;
    %if %sysfunc(exist(_g2, data)) %then %do;
        data _xl."CM3_SameCodeInd_DiffATC"n(dblabel=YES);    set _g2; run;
    %end;

    libname _xl clear;

    /* ---- Cleanup ---- */
    %_dropIfExists(dsn=_ae);   %_dropIfExists(dsn=_cm);
    %_dropIfExists(dsn=_ae1);  %_dropIfExists(dsn=_cm1);
    %_dropIfExists(dsn=_a1);   %_dropIfExists(dsn=_a2);
    %_dropIfExists(dsn=_b1);   %_dropIfExists(dsn=_b2);
    %_dropIfExists(dsn=_c1);   %_dropIfExists(dsn=_c2);
    %_dropIfExists(dsn=_d1);   %_dropIfExists(dsn=_d2);
    %_dropIfExists(dsn=_e1);   %_dropIfExists(dsn=_e2);
    %_dropIfExists(dsn=_f1);   %_dropIfExists(dsn=_f2);
    %_dropIfExists(dsn=_g1);   %_dropIfExists(dsn=_g2);

    %put NOTE: qcCodeCheck — output written to &out\Coding_QC.xls;
%mend qcCodeCheck;
