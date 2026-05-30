/****************************************************************************
 *  Macro: %qcLabCheck
 *  Category: Quality Control
 *  Purpose: Run 4 automated QC checks on laboratory data (LB domain) and
 *           export flagged records to an Excel workbook (Lab_QC.xls).
 *
 *  Checks:
 *    1. Missing unit / normal-range values when result is present
 *    2. Result value far outside normal range (>10x upper or *10 below lower)
 *    3. Same lab test with inconsistent units across visits
 *    4. Non-numeric characters in numeric result fields
 ****************************************************************************/


/* ========== INTERNAL HELPER ========== */

%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;


/* ========== MAIN MACRO ========== */

%macro qcLabCheck(
    lib = ,    /* Path to SAS data library containing the lab dataset */
    dsn = ,    /* Name of the lab dataset (e.g., lb, lb2) */
    out =      /* Output directory for the Lab_QC.xls Excel file */
);

    %local _check;

    /* ---- Validate required parameters ---- */
    %if %length(&lib) = 0 or %length(&dsn) = 0 or %length(&out) = 0 %then %do;
        %put ERROR: qcLabCheck — lib, dsn, and out are required parameters;
        %return;
    %end;

    libname _raw "&lib";

    /* ---- Verify dataset exists ---- */
    %if %sysfunc(exist(_raw.&dsn, data)) = 0 %then %do;
        %put ERROR: qcLabCheck — dataset &dsn not found in &lib;
        libname _raw clear;
        %return;
    %end;

    %put NOTE: qcLabCheck — checking &dsn in &lib;
    %put NOTE: qcLabCheck — output will be written to &out\Lab_QC.xls;

    /* ============================================ */
    /*  Read base lab data                          */
    /* ============================================ */
    data _lb;
        set _raw.&dsn;

        /* Keep core lab variables; use drop/keep for resilience.
           LBSIG may be LBNRIND / LBTOXGR depending on source model. */
        keep usubjid siteid visit form LBTEST LBORRES LBORRESU
             LBORNRLO LBORNRHI LBSIG;
    run;

    /* ============================================ */
    /*  Check 1: Missing unit / normal range        */
    /* ============================================ */
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

    /* ============================================ */
    /*  Check 2: Value range outliers (>10x)        */
    /* ============================================ */
    /* Convert character ranges/results to numeric */
    data _lb2;
        set _lb;
        _NRLO = input(compress(LBORNRLO, '<>='), best.);
        _NRHI = input(compress(LBORNRHI, '<>='), best.);
        _RES  = input(LBORRES, best.);
    run;

    data _fin2;
        set _lb2;
        length _check $100;

        /* Check lower bound: result * 10 < lower limit = suspicious */
        if ^missing(_RES) and ^missing(_NRLO) then do;
            if _RES * 10 < _NRLO then do;
                _check = 'Result *10 below lower limit';
                if ^index(LBSIG, 'Abnormal') then
                    _check = catx('; ', _check, 'not flagged abnormal');
            end;
        end;

        /* Check upper bound: result > 10 * upper limit = outlier */
        if ^missing(_RES) and ^missing(_NRHI) then do;
            if _RES > _NRHI * 10 then
                _check = catx('; ', _check, 'Result >10x upper limit');
        end;

        if _check ^= '' then output;
    run;

    /* ============================================ */
    /*  Check 3: Inconsistent units across visits   */
    /* ============================================ */
    proc sort data=_lb; by LBTEST LBORRESU; run;
    data _fin3;
        set _lb;
        by LBTEST;
        _prevU = lag(LBORRESU);
        if ^first.LBTEST and LBORRESU ^= '' and _prevU ^= ''
           and LBORRESU ^= _prevU then do;
            _check = 'Same test with different units';
            output;
        end;
        drop _prevU;
    run;

    /* ============================================ */
    /*  Check 4: Non-numeric characters in result   */
    /* ============================================ */
    data _fin4;
        set _raw.&dsn;
        keep usubjid visit lbtest lborres LBORNRHI LBORNRLO;
        /* Remove digits and decimal point; whatever remains is non-numeric */
        _trimmed = compress(lborres, '.', 'd');
        if _trimmed ^= '' then output;
    run;

    /* ============================================ */
    /*  Export to Excel                             */
    /* ============================================ */

    libname _xl EXCEL "&out\Lab_QC.xls";

    data _xl."01_Missing_Unit_Range"n(dblabel=YES);
        set _fin1;
    run;
    data _xl."02_Range_Outlier"n(dblabel=YES);
        set _fin2;
    run;
    data _xl."03_Inconsistent_Unit"n(dblabel=YES);
        set _fin3;
    run;
    data _xl."04_NonNumeric_Result"n(dblabel=YES);
        set _fin4;
    run;

    libname _xl clear;

    /* ---- Cleanup ---- */
    %_dropIfExists(dsn=_lb);  %_dropIfExists(dsn=_lb2);
    %_dropIfExists(dsn=_fin1); %_dropIfExists(dsn=_fin2);
    %_dropIfExists(dsn=_fin3); %_dropIfExists(dsn=_fin4);
    libname _raw clear;

    %put NOTE: qcLabCheck — output written to &out\Lab_QC.xls;
%mend qcLabCheck;
