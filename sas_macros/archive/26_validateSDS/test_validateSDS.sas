/****************************************************************************
 *  Test Script: %validateSDS
 *  Purpose: Create mock SDS-like datasets, export them to a temporary
 *           Excel workbook, and run validateSDS to verify rule detection.
 *
 *  Deliberate issues injected (targeting specific rules):
 *    Rule 03: FieldOID "BIRTHDATETIME" (13 chars > 8)
 *    Rule 04: FieldName "Patient Reported Outcome Quality of Life Score" (>40)
 *    Rule 06: Grid field G1 missing Label
 *    Rule 07: Dropdown field DD1 missing DataDictionaryID
 *    Rule 14: FieldOID "SUBJID" vs VariableNo "SUBJECTID" (mismatch)
 *    Rule 18: FormOID "DM_FRM" vs SASText "DM" (mismatch)
 *    Rule 20: Visit "Unscheduled" missing inGroup and outGroup
 *    Rule 24: VariableNo "V001" appears twice
 *
 *  Note: This test creates a temporary Excel file via PROC EXPORT.
 ****************************************************************************/

/* ---- Clean workspace ---- */
proc datasets library=work nolist kill; quit;

%let _tmpDir = %sysfunc(pathname(work));

/* ============================================================ */
/*  DatadictionaryEntry                                          */
/* ============================================================ */
/* Columns: dataDictionaryOID, entryOID, ordinal, itemDataString,
            isSpecify, isActive, clinicalType, labCommentAvailable,
            clinicalQueryString */
data _dic_raw;
    length dataDictionaryOID $20 entryOID $20 ordinal 8
           itemDataString $40 isSpecify $4 isActive $4
           clinicalType $20 labCommentAvailable $4
           clinicalQueryString $40;
    /* Row 1: Normal entry */
    dataDictionaryOID = "SEX"; entryOID = "1"; ordinal = 1;
    itemDataString = "Male"; isSpecify = "0"; isActive = "1";
    clinicalType = ""; labCommentAvailable = "0"; output;
    /* Row 2: Normal entry */
    dataDictionaryOID = "SEX"; entryOID = "2"; ordinal = 2;
    itemDataString = "Female"; isSpecify = "0"; isActive = "1";
    clinicalType = ""; labCommentAvailable = "0"; output;
    /* Row 3: Clinical dictionary entry missing clinicalType */
    dataDictionaryOID = "AE_TERM"; entryOID = "1"; ordinal = 1;
    itemDataString = "MedDRA Lookup"; isSpecify = "0"; isActive = "1";
    clinicalType = ""; clinicalQueryString = "";
    labCommentAvailable = "0"; output;
run;

/* ============================================================ */
/*  DataDictionary                                               */
/* ============================================================ */
/* Columns: dataDictionaryOID, isClinical */
data _ic_raw;
    length dataDictionaryOID $20 isClinical $4;
    dataDictionaryOID = "SEX";     isClinical = "0"; output;
    dataDictionaryOID = "AE_TERM"; isClinical = "1"; output;
    dataDictionaryOID = "RACE";    isClinical = "0"; output;
run;

/* ============================================================ */
/*  Field (main SDS sheet)                                       */
/* ============================================================ */
data _field_raw;
    length fieldOID $40 fieldName $80 controlType $30
           label $20 unit $20 requireVerification $4
           reviewGroups $40 dataFormat $40 dataDictionaryOID $20
           isLab $4 variableNo $40 isFutureDate $4
           formOID $20 SASText $20 labKey $40
           gridDefaultValueDictionary $10 isGridCanAddRow $4
           isClinicalRequired $4 viewRestrictions $40 isRequireSign $4;

    /* Row 1: Normal field — no issues */
    fieldOID = "SUBJID"; fieldName = "Subject ID";
    controlType = "Text box"; label = "Subject";
    unit = ""; requireVerification = "1";
    reviewGroups = "DM Review"; dataFormat = "";
    dataDictionaryOID = ""; isLab = "0";
    variableNo = "SUBJID"; isFutureDate = "0";
    formOID = "DM"; SASText = "DM"; labKey = "";
    gridDefaultValueDictionary = ""; isGridCanAddRow = "0";
    isClinicalRequired = "0"; viewRestrictions = "DM";
    isRequireSign = "0"; output;

    /* Row 2: Rule 03 — FieldOID > 8 chars */
    fieldOID = "BIRTHDATETIME"; fieldName = "Birth Date Time";
    controlType = "Text box"; label = "Birth DT";
    unit = ""; requireVerification = "1";
    reviewGroups = "DM Review"; dataFormat = "";
    dataDictionaryOID = ""; isLab = "0";
    variableNo = "BIRTHDT"; isFutureDate = "0";
    formOID = "DM"; SASText = "DM"; labKey = "";
    gridDefaultValueDictionary = ""; isGridCanAddRow = "0";
    isClinicalRequired = "0"; viewRestrictions = "";
    isRequireSign = "0"; output;

    /* Row 3: Rule 04 — FieldName > 40 chars */
    fieldOID = "PROQOL"; fieldName = "Patient Reported Outcome Quality of Life Score Form v2";
    controlType = "Text box"; label = "PRO";
    unit = ""; requireVerification = "1";
    reviewGroups = "DM Review"; dataFormat = "";
    dataDictionaryOID = ""; isLab = "0";
    variableNo = "PROQOL"; isFutureDate = "0";
    formOID = "QOL"; SASText = "QOL"; labKey = "";
    gridDefaultValueDictionary = ""; isGridCanAddRow = "0";
    isClinicalRequired = "0"; viewRestrictions = "";
    isRequireSign = "0"; output;

    /* Row 4: Rule 06 — Grid field no label */
    fieldOID = "G1"; fieldName = "Grid Field 1";
    controlType = "Grid"; label = "";
    unit = ""; requireVerification = "1";
    reviewGroups = "DM Review"; dataFormat = "";
    dataDictionaryOID = ""; isLab = "0";
    variableNo = "G1"; isFutureDate = "0";
    formOID = "CONMED"; SASText = "CONMED"; labKey = "";
    gridDefaultValueDictionary = ""; isGridCanAddRow = "0";
    isClinicalRequired = "0"; viewRestrictions = "";
    isRequireSign = "0"; output;

    /* Row 5: Rule 07 — Dropdown missing DictID */
    fieldOID = "DD1"; fieldName = "DropDown Test";
    controlType = "Dropdown"; label = "Select One";
    unit = ""; requireVerification = "1";
    reviewGroups = "DM Review"; dataFormat = "";
    dataDictionaryOID = ""; isLab = "0";
    variableNo = "DD1"; isFutureDate = "0";
    formOID = "AE"; SASText = "AE"; labKey = "";
    gridDefaultValueDictionary = ""; isGridCanAddRow = "0";
    isClinicalRequired = "0"; viewRestrictions = "";
    isRequireSign = "0"; output;

    /* Row 6: Rule 14 — FieldOID != VariableNo */
    fieldOID = "SUBJID2"; fieldName = "Subject ID Alt";
    controlType = "Text box"; label = "Subj ID";
    unit = ""; requireVerification = "1";
    reviewGroups = "DM Review"; dataFormat = "";
    dataDictionaryOID = ""; isLab = "0";
    variableNo = "SUBJECTID"; isFutureDate = "0";
    formOID = "DM"; SASText = "DM"; labKey = "";
    gridDefaultValueDictionary = ""; isGridCanAddRow = "0";
    isClinicalRequired = "0"; viewRestrictions = "";
    isRequireSign = "0"; output;

    /* Row 7: Rule 18 — FormOID != SASText */
    fieldOID = "AETERM"; fieldName = "AE Term";
    controlType = "Text box"; label = "AE Name";
    unit = ""; requireVerification = "1";
    reviewGroups = "DM Review"; dataFormat = "";
    dataDictionaryOID = ""; isLab = "0";
    variableNo = "AETERM"; isFutureDate = "0";
    formOID = "AE_FRM"; SASText = "AE"; labKey = "";
    gridDefaultValueDictionary = ""; isGridCanAddRow = "0";
    isClinicalRequired = "0"; viewRestrictions = "";
    isRequireSign = "0"; output;

    /* Row 8: Rule 24 — Duplicate VariableNo (same as row 6) */
    fieldOID = "FIELDXX"; fieldName = "Extra Field";
    controlType = "Text box"; label = "Extra";
    unit = ""; requireVerification = "1";
    reviewGroups = "DM Review"; dataFormat = "";
    dataDictionaryOID = ""; isLab = "0";
    variableNo = "SUBJECTID"; isFutureDate = "0";
    formOID = "DM"; SASText = "DM"; labKey = "";
    gridDefaultValueDictionary = ""; isGridCanAddRow = "0";
    isClinicalRequired = "0"; viewRestrictions = "";
    isRequireSign = "0"; output;

    /* Row 9: Lab field without isClinicalRequired */
    fieldOID = "LABTEST1"; fieldName = "Lab Test 1";
    controlType = "Lab"; label = "Lab1";
    unit = ""; requireVerification = "1";
    reviewGroups = "DM Review"; dataFormat = "";
    dataDictionaryOID = ""; isLab = "1";
    variableNo = "LABTEST1"; isFutureDate = "0";
    formOID = "LB"; SASText = "LB"; labKey = "ALT";
    gridDefaultValueDictionary = ""; isGridCanAddRow = "0";
    isClinicalRequired = "0"; viewRestrictions = "";
    isRequireSign = "0"; output;

    /* Row 10: Normal — no issues */
    fieldOID = "WEIGHT"; fieldName = "Weight";
    controlType = "Text box"; label = "Weight";
    unit = "kg"; requireVerification = "1";
    reviewGroups = "DM Review"; dataFormat = "";
    dataDictionaryOID = ""; isLab = "0";
    variableNo = "WEIGHT"; isFutureDate = "0";
    formOID = "VS"; SASText = "VS"; labKey = "";
    gridDefaultValueDictionary = ""; isGridCanAddRow = "0";
    isClinicalRequired = "0"; viewRestrictions = "";
    isRequireSign = "0"; output;
run;

/* ============================================================ */
/*  VisitWindowSetting                                           */
/* ============================================================ */
data _vws_raw;
    length visitName $20 inGroup $10 outGroup $10;
    /* Row 1: Normal */
    visitName = "Week 0"; inGroup = "0"; outGroup = "7"; output;
    /* Row 2: Normal */
    visitName = "Week 4"; inGroup = "-3"; outGroup = "3"; output;
    /* Row 3: Rule 20 — Missing inGroup and outGroup */
    visitName = "Unscheduled"; inGroup = ""; outGroup = ""; output;
    /* Row 4: Rule 20 — Missing outGroup */
    visitName = "Week 8"; inGroup = "-3"; outGroup = ""; output;
run;

/* ============================================================ */
/*  LabKey reference (optional)                                  */
/* ============================================================ */
data _lab_raw;
    length LabKey $40 KeyDescription $80;
    LabKey = "ALT"; KeyDescription = "Alanine Aminotransferase"; output;
    LabKey = "AST"; KeyDescription = "Aspartate Aminotransferase"; output;
    LabKey = "GLUC"; KeyDescription = "Glucose"; output;
run;

/* ---- Export all to Excel ---- */
proc export data=_dic_raw
    outfile="&_tmpDir\_test_SDS.xlsx"
    dbms=xlsx replace;
    sheet="DatadictionaryEntry";
run;
proc export data=_ic_raw
    outfile="&_tmpDir\_test_SDS.xlsx"
    dbms=xlsx replace;
    sheet="DataDictionary";
run;
proc export data=_field_raw
    outfile="&_tmpDir\_test_SDS.xlsx"
    dbms=xlsx replace;
    sheet="Field";
run;
proc export data=_vws_raw
    outfile="&_tmpDir\_test_SDS.xlsx"
    dbms=xlsx replace;
    sheet="VisitWindowSetting";
run;

proc export data=_lab_raw
    outfile="&_tmpDir\_test_labKey.xlsx"
    dbms=xlsx replace;
    sheet="LabKey";
run;

/* ---- Run %validateSDS ---- */
title "Test: %validateSDS — 25 Rule-Based SDS Checks";

%validateSDS(
    sds    = &_tmpDir\_test_SDS.xlsx,
    labKey = &_tmpDir\_test_labKey.xlsx,
    out    = &_tmpDir
);

/* ---- Verify key results ---- */

title2 "Rule 03: FieldOID > 8 chars (expect 1: BIRTHDATETIME)";
%if %sysfunc(exist(_f03, data)) %then %do;
    proc print data=_f03 noobs; var fieldOID fieldName; run;
%end;

title2 "Rule 04: FieldName > 40 chars (expect 1: PROQOL)";
%if %sysfunc(exist(_f04, data)) %then %do;
    proc print data=_f04 noobs; var fieldOID fieldName; run;
%end;

title2 "Rule 06: Grid no label (expect 1: G1)";
%if %sysfunc(exist(_f06, data)) %then %do;
    proc print data=_f06 noobs; var fieldOID fieldName label; run;
%end;

title2 "Rule 07: Dropdown no DictID (expect 1: DD1)";
%if %sysfunc(exist(_f07, data)) %then %do;
    proc print data=_f07 noobs; var fieldOID controlType dataDictionaryOID; run;
%end;

title2 "Rule 14: FieldOID != VariableNo (expect 1: SUBJID2)";
%if %sysfunc(exist(_f14, data)) %then %do;
    proc print data=_f14 noobs; var fieldOID variableNo; run;
%end;

title2 "Rule 18: FormOID != SASText (expect 1: AETERM)";
%if %sysfunc(exist(_f18, data)) %then %do;
    proc print data=_f18 noobs; var fieldOID formOID SASText; run;
%end;

title2 "Rule 20: Missing in/out group (expect 2: Unsched, Week8)";
%if %sysfunc(exist(_f20, data)) %then %do;
    proc print data=_f20 noobs; var visitName inGroup outGroup; run;
%end;

title2 "Rule 24: Duplicate VariableNo (expect 2: SUBJECTID)";
%if %sysfunc(exist(_f24, data)) %then %do;
    proc print data=_f24 noobs; var fieldOID variableNo; run;
%end;

title2 "Rule 25: Lab not ClinicalRequired (expect 1: LABTEST1)";
%if %sysfunc(exist(_f25, data)) %then %do;
    proc print data=_f25 noobs; var fieldOID isLab isClinicalRequired; run;
%end;

title "Test complete.";
