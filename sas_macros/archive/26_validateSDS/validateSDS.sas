/****************************************************************************
 *  Macro: %validateSDS
 *  Category: Quality Control
 *  Purpose: Comprehensive SDS (Study Design Specification) validation with
 *           25 rule-based checks across DatadictionaryEntry, DataDictionary,
 *           Field, VisitWindowSetting, and optionally LabKey sheets.
 *
 *  Reads all key sheets from an SDS Excel workbook and flags violations.
 *  Results are written to a timestamped Excel workbook
 *  (SDS_Validation_YYYY-MM-DD.xlsx), one sheet per rule.
 *
 *  Rules 01-25 (detailed in parameter section below).
 ****************************************************************************/


/* ========== INTERNAL HELPER ========== */

%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;


/* ========== MAIN MACRO ========== */

%macro validateSDS(
    sds    = ,   /* Full path to the SDS Excel workbook */
    labKey = ,   /* Full path to the LabKey reference Excel file (optional) */
    out    =     /* Output directory for the validation report */
);

    %local _rundate;

    /* ---- Validate required parameters ---- */
    %if %length(&sds) = 0 or %length(&out) = 0 %then %do;
        %put ERROR: validateSDS — sds and out are required parameters;
        %return;
    %end;

    /* ---- Verify SDS file exists ---- */
    %if %sysfunc(fileexist("&sds")) = 0 %then %do;
        %put ERROR: validateSDS — SDS file not found: &sds;
        %return;
    %end;

    %let _rundate = %sysfunc(date(), yymmdd10.);
    %put NOTE: validateSDS — processing &sds;
    %put NOTE: validateSDS — output will be written to &out\SDS_Validation_&_rundate..xlsx;

    /* ============================================ */
    /*  Import all SDS sheets                       */
    /* ============================================ */

    proc import datafile="&sds" out=_dic dbms=xlsx replace;
        sheet="DatadictionaryEntry"; getnames=yes;
    run;
    proc import datafile="&sds" out=_ic dbms=xlsx replace;
        sheet="DataDictionary"; getnames=yes;
    run;
    proc import datafile="&sds" out=_field dbms=xlsx replace;
        sheet="Field"; getnames=yes;
    run;
    proc import datafile="&sds" out=_vws dbms=xlsx replace;
        sheet="VisitWindowSetting"; getnames=yes;
    run;

    /* ---- Optional: import LabKey reference ---- */
    %if %length(&labKey) > 0 %then %do;
        %if %sysfunc(fileexist("&labKey")) %then %do;
            proc import datafile="&labKey" out=_lab dbms=xlsx replace;
                sheet="LabKey"; getnames=yes;
            run;
        %end;
        %else %do;
            %put WARNING: validateSDS — LabKey file not found: &labKey, skipping LabKey checks;
        %end;
    %end;

    /* ====================================================== */
    /*  RULE 01: Duplicate ItemDataString rows (active only)  */
    /* ====================================================== */
    data _dic1;
        set _dic;
        if isActive = '1';
    run;
    proc sort data=_dic1; by dataDictionaryOID entryOID; run;
    data _f01;
        set _dic1;
        if itemDataString ^= '';
        _l1 = lag(dataDictionaryOID); _l2 = lag(entryOID);
        _l3 = lag(ordinal);           _l4 = lag(itemDataString);
        _l5 = lag(isSpecify);         _l6 = lag(isActive);
        _l7 = lag(clinicalType);      _l8 = lag(labCommentAvailable);
        if _l1 = dataDictionaryOID and _l2 = entryOID and _l3 = ordinal
           and _l4 = itemDataString and _l5 = isSpecify and _l6 = isActive
           and _l7 = clinicalType and _l8 = labCommentAvailable
           and ^first.entryOID;
        drop _l1-_l8;
    run;

    /* ====================================================== */
    /*  RULE 02: ClinicalType / ClinicalQueryString empty     */
    /* ====================================================== */
    data _ic2;
        set _ic;
        if isClinical = '1';
        keep dataDictionaryOID;
    run;

    data _f02;
        merge _dic(in=_a) _ic2(in=_b);
        by dataDictionaryOID;
        if _a and _b and (clinicalType = '' or clinicalQueryString = '');
    run;

    /* ====================================================== */
    /*  RULE 03: FieldOID longer than 8 characters           */
    /* ====================================================== */
    data _f03;
        set _field;
        if length(fieldOID) > 8;
    run;

    /* ====================================================== */
    /*  RULE 04: FieldName longer than 40 characters         */
    /* ====================================================== */
    data _f04;
        set _field;
        if length(fieldName) > 40;
    run;

    /* ====================================================== */
    /*  RULE 05: Date control field missing Unit             */
    /* ====================================================== */
    data _f05;
        set _field;
        if fieldName = 'Date control' and unit = '';
    run;

    /* ====================================================== */
    /*  RULE 06: Grid field missing Label                    */
    /* ====================================================== */
    data _f06;
        set _field;
        where label = '';
    run;

    /* ====================================================== */
    /*  RULE 07: Radio/dropdown/checkbox missing DictID       */
    /* ====================================================== */
    data _f07;
        set _field;
        if controlType in ('Horizontal radio','Vertical radio','Dropdown')
           and dataDictionaryOID = '';
    run;

    /* ====================================================== */
    /*  RULE 08: Text-label has unexpected ReviewGroups       */
    /* ====================================================== */
    data _f08;
        set _field;
        if controlType = 'Text label' and reviewGroups ^= '';
    run;

    /* ====================================================== */
    /*  RULE 09: Text-label has RequireVerification = 1       */
    /* ====================================================== */
    data _f09;
        set _field;
        if controlType = 'Text label' and requireVerification = '1';
    run;

    /* ====================================================== */
    /*  RULE 10: Non-select non-lab has DataDictionaryID     */
    /* ====================================================== */
    data _f10;
        set _field;
        if controlType not in ('Horizontal radio','Vertical radio','Dropdown','Lab')
           and isLab = '0' and dataDictionaryOID ^= '';
    run;

    /* ====================================================== */
    /*  RULE 11: Non-text-label missing ReviewGroups          */
    /* ====================================================== */
    data _f11;
        set _field;
        if controlType ^= 'Text label' and reviewGroups = '';
    run;

    /* ====================================================== */
    /*  RULE 12: Non-text-label RequireVerification = 0       */
    /* ====================================================== */
    data _f12;
        set _field;
        if controlType ^= 'Text label' and requireVerification = '0';
    run;

    /* ====================================================== */
    /*  RULE 13: Date (yyyy-MM-dd) missing IsFutureDateTime=1 */
    /* ====================================================== */
    data _f13;
        set _field;
        if index(dataFormat, 'yyyy-MM-dd') and isFutureDate ^= '1';
    run;

    /* ====================================================== */
    /*  RULE 14: FieldOID not equal to VariableNo             */
    /* ====================================================== */
    data _f14;
        set _field;
        if controlType ^= 'Text label' and fieldOID ^= variableNo;
    run;

    /* ====================================================== */
    /*  RULE 15: FieldOID not equal to labKey                 */
    /* ====================================================== */
    data _f15;
        set _field;
        where fieldOID ^= labKey and labKey ^= '';
    run;

    /* ====================================================== */
    /*  RULE 16: Duplicate labKey usage                       */
    /* ====================================================== */
    proc sort data=_field(where=(labKey ^= '')) nouniquekey out=_f16;
        by labKey;
    run;

    /* ====================================================== */
    /*  RULE 17: labKey description mismatch (fieldName)      */
    /* ====================================================== */
    %if %sysfunc(exist(_lab, data)) %then %do;
        data _lb1;
            set _field;
            where labKey ^= '';
            keep formOID fieldOID fieldName labKey;
        run;
        proc sort data=_lb1; by labKey; run;
        proc sort data=_lab(keep=LabKey KeyDescription) out=_lb2;
            by labKey;
        run;
        data _f17;
            merge _lb1(in=_a) _lb2;
            by labKey;
            if _a and fieldName ^= KeyDescription;
        run;
    %end;

    /* ====================================================== */
    /*  RULE 18: FormOID not equal to SASText                 */
    /* ====================================================== */
    data _f18;
        set _field;
        where formOID ^= SASText;
    run;

    /* ====================================================== */
    /*  RULE 19: Grid default value with row-add enabled      */
    /* ====================================================== */
    data _f19;
        set _field;
        if gridDefaultValueDictionary ^= '' and isGridCanAddRow = '1';
    run;

    /* ====================================================== */
    /*  RULE 20: Visit window missing inGroup or outGroup     */
    /* ====================================================== */
    data _f20;
        set _vws;
        where inGroup = '' or outGroup = '';
    run;

    /* ====================================================== */
    /*  RULE 21: DM view restriction without DM Review group  */
    /* ====================================================== */
    data _f21;
        set _field;
        if index(viewRestrictions, 'DM') and ^index(reviewGroups, 'DM Review');
    run;

    /* ====================================================== */
    /*  RULE 22: CRA view restriction without RequireVerify   */
    /* ====================================================== */
    data _f22;
        set _field;
        if index(viewRestrictions, 'CRA') and requireVerification = '1';
    run;

    /* ====================================================== */
    /*  RULE 23: PI view restriction without IsRequireSign    */
    /* ====================================================== */
    data _f23;
        set _field;
        if index(viewRestrictions, 'PI') and isRequireSign = '1';
    run;

    /* ====================================================== */
    /*  RULE 24: Duplicate VariableNo                         */
    /* ====================================================== */
    proc sort data=_field(where=(variableNo ^= '')) nouniquekey out=_f24;
        by variableNo;
    run;

    /* ====================================================== */
    /*  RULE 25: Lab field without IsClinicalRequired         */
    /* ====================================================== */
    data _f25;
        set _field;
        if isLab = '1' and isClinicalRequired = '0';
    run;

    /* ============================================ */
    /*  Export all results to Excel                 */
    /* ============================================ */

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

    %if %sysfunc(exist(_f17, data)) %then %do;
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

    /* ============================================ */
    /*  Cleanup                                     */
    /* ============================================ */

    %_dropIfExists(dsn=_dic);   %_dropIfExists(dsn=_ic);
    %_dropIfExists(dsn=_field); %_dropIfExists(dsn=_vws);
    %_dropIfExists(dsn=_lab);   %_dropIfExists(dsn=_dic1);
    %_dropIfExists(dsn=_ic2);
    %_dropIfExists(dsn=_f01);  %_dropIfExists(dsn=_f02);
    %_dropIfExists(dsn=_f03);  %_dropIfExists(dsn=_f04);
    %_dropIfExists(dsn=_f05);  %_dropIfExists(dsn=_f06);
    %_dropIfExists(dsn=_f07);  %_dropIfExists(dsn=_f08);
    %_dropIfExists(dsn=_f09);  %_dropIfExists(dsn=_f10);
    %_dropIfExists(dsn=_f11);  %_dropIfExists(dsn=_f12);
    %_dropIfExists(dsn=_f13);  %_dropIfExists(dsn=_f14);
    %_dropIfExists(dsn=_f15);  %_dropIfExists(dsn=_f16);
    %_dropIfExists(dsn=_f17);
    %_dropIfExists(dsn=_f18);  %_dropIfExists(dsn=_f19);
    %_dropIfExists(dsn=_f20);  %_dropIfExists(dsn=_f21);
    %_dropIfExists(dsn=_f22);  %_dropIfExists(dsn=_f23);
    %_dropIfExists(dsn=_f24);  %_dropIfExists(dsn=_f25);
    %_dropIfExists(dsn=_lb1);  %_dropIfExists(dsn=_lb2);

    %put NOTE: validateSDS — output written to &out\SDS_Validation_&_rundate..xlsx;
%mend validateSDS;
