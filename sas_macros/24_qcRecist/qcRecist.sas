/****************************************************************************
 *  Macro: %qcRecist
 *  Category: Quality Control
 *  Purpose: Verify RECIST 1.1 tumor response evaluation consistency.
 *           Compares EDC-entered target response (CR/PR/SD/PD) against
 *           calculated values from target-lesion measurements.
 *
 *  Two reports are exported:
 *    Sheet "Recist"       — Target response comparison (EDC vs calculated)
 *    Sheet "Consistency"  — Lesion-level method/location consistency
 *
 *  Calculated responses:
 *    PD — Sum of Diameters > smallest previous sum * 1.2 AND absolute
 *          increase > 5 mm
 *    SD — Neither PD nor PR criteria met
 *    PR — Sum of Diameters <= baseline sum * 0.7
 *    CR — All lesions disappeared (non-lymph < 10mm or absent)
 ****************************************************************************/


/* ========== INTERNAL HELPER ========== */

%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;


/* ========== MAIN MACRO ========== */

%macro qcRecist(
    lib      = ,      /* Path to SAS data library */
    out      = ,      /* Output directory for Excel reports */
    scrDS    = ,      /* Screening visit dataset name */
    fuDS     = ,      /* Follow-up visit dataset name */
    respDS   = ,      /* Dataset with EDC-assessed overall response */
    respVar  = ,      /* EDC response variable in respDS (e.g., TRGRESP) */
    scrDateVar = ,    /* Screening assessment date variable */
    fuDateVar  = ,    /* Follow-up assessment date variable */
    scrLongVar = ,    /* Screening longest diameter variable */
    fuLongVar  = ,    /* Follow-up longest diameter variable */
    scrSumVar  = ,    /* Screening sum of diameters variable */
    fuSumVar   = ,    /* Follow-up sum of diameters variable */
    scrLocVar  = ,    /* Screening lesion location variable */
    fuLocVar   = ,    /* Follow-up lesion location variable */
    scrSiteVar = ,    /* Screening location-within-site variable */
    fuSiteVar  = ,    /* Follow-up location-within-site variable */
    scrRecVar  = ,    /* Screening record position variable */
    fuRecVar   = ,    /* Follow-up record position variable */
    scrMetVar  = ,    /* Screening method variable */
    fuMetVar   = ,    /* Follow-up method variable */
    scrMetSpecVar = , /* Screening method-specify variable */
    fuMetSpecVar  = , /* Follow-up method-specify variable */
    scrLabel   = Screening   /* Screening visit label */
);

    %local _rundate _xlsxname;

    /* ---- Validate required parameters ---- */
    %if %length(&lib) = 0 or %length(&out) = 0 %then %do;
        %put ERROR: qcRecist — lib and out are required parameters;
        %return;
    %end;
    %if %length(&scrDS) = 0 or %length(&fuDS) = 0 or %length(&respDS) = 0 %then %do;
        %put ERROR: qcRecist — scrDS, fuDS, and respDS are required;
        %return;
    %end;
    %if %length(&respVar) = 0 %then %do;
        %put ERROR: qcRecist — respVar is required;
        %return;
    %end;

    libname _recist "&lib";

    %let _rundate  = %sysfunc(date(), yymmddn8.);
    %let _xlsxname = Recist&_rundate.;

    %put NOTE: qcRecist — checking RECIST in &lib;
    %put NOTE: qcRecist — screening ds=&scrDS, follow-up ds=&fuDS, response ds=&respDS;

    /* ================================================================ */
    /*  PART A: Target Response Calculation vs EDC                       */
    /* ================================================================ */

    /* ---- Import screening and follow-up measurement data ---- */
    data _rs_scr;
        set _recist.&scrDS;
        rename &scrDateVar = TRCDAT1
               &scrLongVar = LDIAM
               &scrSumVar  = SLDIAM
               &scrLocVar  = TULOC1
               &scrSiteVar = S_R1SIT
               &scrRecVar  = RecordPosition;
        keep Subject InstanceName &scrDateVar &scrSumVar &scrLongVar
             &scrLocVar &scrSiteVar &scrRecVar;
    run;

    data _rs_fu;
        set _recist.&fuDS;
        rename &fuDateVar = TRCDAT1
               &fuLongVar = LDIAM
               &fuSumVar  = SLDIAM
               &fuLocVar  = TULOC1
               &fuSiteVar = S_R1SIT
               &fuRecVar  = RecordPosition;
        keep Subject InstanceName &fuDateVar &fuSumVar &fuLongVar
             &fuLocVar &fuSiteVar &fuRecVar;
    run;

    data _rs_all;
        set _rs_scr _rs_fu;
    run;

    /* ---- Harmonize lesion identifiers ---- */
    data _rs_all;
        set _rs_all;
        z1 = cats(TULOC1, S_R1SIT);
        a  = input(compress(TRCDAT1), date9.);
        if a = . then a = input(compress(TRCDAT1), date8.);
    run;

    /* ---- Find smallest previous SLDIAM per subject ---- */
    proc sort data=_rs_all; by Subject InstanceName; run;
    proc sort data=_rs_all; by Subject a RecordPosition; run;

    data _rs_rank;
        set _rs_all;
        retain minSLDIAM;
        if first.Subject then minSLDIAM = SLDIAM;
        else minSLDIAM = min(SLDIAM, minSLDIAM);
        by Subject;
        cz = SLDIAM / minSLDIAM;
    run;

    /* ---- Get screening visit data for baseline references ---- */
    data _scr_ref;
        set _rs_all;
        z2 = 1;
        if InstanceName = "&scrLabel";
        keep Subject z1 z2;
    run;

    data _scr_base;
        set _rs_rank;
        if InstanceName = "&scrLabel";
        keep Subject SLDIAM;
        rename SLDIAM = SLDIAMa;
    run;

    proc sort data=_scr_ref;  by Subject z1; run;
    proc sort data=_scr_base; by Subject; run;
    proc sort data=_rs_rank;  by Subject z1; run;

    data _rs_rank;
        merge _rs_rank _scr_ref;
        by Subject z1;
    run;
    data _rs_rank;
        merge _rs_rank _scr_base;
        by Subject;
    run;

    /* ---- Calculate target responses ---- */
    data _rs_calc;
        set _rs_rank;
        cz1   = SLDIAM / SLDIAMa;
        result  = '';   /* PD */
        result1 = '';   /* PR */
        result2 = '';   /* SD */
        result3 = '';   /* CR */

        /* PD: >20% increase AND >5mm absolute increase */
        if cz > 1.2 and abs(SLDIAM - minSLDIAM) > 5 then result = 'PD';

        /* SD: not PD and not PR */
        if cz < 1.2 and cz1 > 0.7 and SLDIAM ^= '' and SLDIAMa ^= ''
           and InstanceName ^= "&scrLabel" then result2 = 'SD';
        if abs(SLDIAM - minSLDIAM) < 5 and cz1 > 0.7 and SLDIAM ^= ''
           and SLDIAMa ^= '' and InstanceName ^= "&scrLabel" then result2 = 'SD';

        /* PR: >=30% decrease from baseline */
        if cz1 <= 0.7 and SLDIAM ^= '' and SLDIAMa ^= '' then result1 = 'PR';

        q4 = SLDIAM - minSLDIAM;

        label q4     = 'Difference: current SLD minus smallest previous SLD';
        label result2 = 'Calculated target response (SD)';
        label result1 = 'Calculated target response (PR)';
        label result  = 'Calculated target response (PD)';
        label cz     = 'Ratio: current SLD / smallest previous SLD';
        label cz1    = 'Ratio: current SLD / baseline SLD';
        label minSLDIAM = 'Smallest previous Sum of Diameters';
    run;

    /* ---- CR check: lesion disappearance ---- */
    /* Find lesions that disappear between visits */
    data _cr_chk;
        set _rs_calc;
        keep Subject TRCDAT1 TULOC1 InstanceName;
        rename TRCDAT1 = TRCDAT1cr TULOC1 = TULOC1cr InstanceName = InstanceNamecr;
    run;

    proc sort data=_rs_calc; by Subject TRCDAT1; run;
    proc sort data=_cr_chk;  by Subject TRCDAT1cr; run;

    proc sql;
        create table _cr_x as
            select * from _cr_chk as cr
            inner join _rs_calc as f
            on cr.Subject = f.Subject;
    quit;

    data _cr_x;
        set _cr_x;
        TRCDAT1crs = input(compress(TRCDAT1cr), date8.);
        TRCDAT1s   = input(compress(TRCDAT1),   date8.);
        if TRCDAT1crs = . then TRCDAT1crs = input(compress(TRCDAT1cr), date9.);
        if TRCDAT1s   = . then TRCDAT1s   = input(compress(TRCDAT1),   date9.);
    run;

    data _cr_x;
        set _cr_x;
        if TRCDAT1crs < TRCDAT1s and TULOC1cr = TULOC1 then YN = 'Yes';
    run;

    data _cr_disappear;
        set _cr_x;
        keep Subject InstanceNamecr TULOC1cr InstanceName TRCDAT1cr;
        if YN = 'Yes';
    run;

    data _cr_flag;
        set _cr_disappear;
        l = 1;
        rename TULOC1cr = TULOC1;
        keep Subject InstanceName TULOC1cr l;
    run;

    proc sort data=_cr_flag; by Subject InstanceName TULOC1; run;
    proc sort data=_rs_calc nodupkey; by Subject InstanceName TRCDAT1; run;

    /* ---- Clean up target response dataset for merge ---- */
    data _rs_clean;
        set _rs_calc;
        drop SLDIAMa LDIAM RecordPosition a z1 z2 TULOC1 S_R1SIT;
        if TRCDAT1 ^= '';
    run;

    /* ---- Import EDC-assessed response ---- */
    data _edc_resp;
        set _recist.&respDS;
        rename &respVar = TRGRESP;
        keep Subject InstanceName &respVar;
    run;

    proc sort data=_rs_clean; by Subject InstanceName; run;
    proc sort data=_edc_resp; by Subject InstanceName; run;

    /* ---- Compare EDC vs calculated responses ---- */
    data _rs_compare;
        length q $200;
        merge _rs_clean _edc_resp(in=x);
        if not x then TRGRESP = "-";
        by Subject InstanceName;

        /* PD inconsistency */
        if TRGRESP = 'PD' and result  ^= 'PD' then q = 'Inconsistent';
        if TRGRESP ^= 'PD' and result  = 'PD' then q = 'Inconsistent';

        /* PR inconsistency */
        if TRGRESP = 'PR' and result1 ^= 'PR' and result ^= 'PD' then q = 'Inconsistent';
        if TRGRESP ^= 'PR' and result1 = 'PR' and result ^= 'PD' then q = 'Inconsistent';

        /* SD inconsistency */
        if TRGRESP = 'SD' and result2 ^= 'SD' then q = 'Inconsistent';
        if TRGRESP ^= 'SD' and result2 = 'SD' then q = 'Inconsistent';

        /* Multiple standards */
        if result1 ^= '' and result  ^= '' then q = 'Multiple standards';
        if result1 ^= '' and result2 ^= '' then q = 'Multiple standards';
        if result2 ^= '' and result  ^= '' then q = 'Multiple standards';

        /* Remove screening-only records without measurements */
        if TRCDAT1 = '' and SLDIAM = '' then delete;

        label q = 'Consistency Result';
        label TRGRESP = 'EDC Target Response';
    run;

    /* ---- CR lesion-type check ---- */
    proc sort data=_rs_calc; by Subject InstanceName TULOC1; run;

    data _cr_merge;
        merge _rs_calc _cr_flag;
        by Subject InstanceName TULOC1;
        label l = 'Lesion not disappeared';
    run;

    data _cr_classify;
        set _cr_merge;
        length l1 $200;
        if TULOC1 ^= '' and InstanceName ^= "&scrLabel";
        if l = '' then l1 = 'YES';
        keep Subject InstanceName l1;
        if l1 ^= '';
        label l1 = 'Is there any disappearance of lesions during this visit';
    run;

    /* Classify lesion types per visit */
    proc sort data=_rs_all out=_cr_loc nodupkey;
        by Subject InstanceName TULOC1;
    run;

    data _cr_loc;
        set _cr_loc;
        fa = cats(Subject, InstanceName);
        keep fa Subject InstanceName TULOC1;
    run;

    proc sort data=_cr_loc; by fa; run;

    data _cr_loc;
        set _cr_loc;
        TULOC2 = lag(TULOC1);
        if first.fa then TULOC2 = '';
        by fa;
    run;

    data _cr_ln _cr_nonln _cr_both;
        set _cr_loc;
        length fb1 fb2 fb3 $200;

        if index(TULOC2, 'Lymph Node') and TULOC2 ^= '' then output _cr_both;
        if index(TULOC1, 'Lymph Node') and TULOC2 ^= '' then output _cr_both;

        if index(TULOC1, 'Lymph Node') and TULOC2 = '' then output _cr_ln;
        if index(TULOC1, 'Lymph Node') = 0 and TULOC2 = '' then output _cr_nonln;
    run;

    data _cr_both;
        set _cr_both;
        fb1 = 'Both non lymph node and lymph node';
        keep Subject InstanceName fb1;
    run;
    data _cr_ln;
        set _cr_ln;
        fb2 = 'Only lymph node';
        keep Subject InstanceName fb2;
    run;
    data _cr_nonln;
        set _cr_nonln;
        fb3 = 'Only non lymph node';
        keep Subject InstanceName fb3;
    run;

    proc sort data=_cr_both  nodupkey; by Subject InstanceName; run;
    proc sort data=_cr_ln    nodupkey; by Subject InstanceName; run;
    proc sort data=_cr_nonln nodupkey; by Subject InstanceName; run;

    data _cr_type;
        merge _cr_both _cr_ln _cr_nonln;
        by Subject InstanceName;
        fb = fb1;
        if fb = '' then fb = fb2;
        if fb = '' then fb = fb3;
        drop fb1 fb2 fb3;
        label fb = 'Type of lesion in this visit';
    run;

    proc sort data=_cr_type;   by Subject InstanceName; run;
    proc sort data=_cr_classify; by Subject InstanceName; run;
    proc sort data=_rs_compare;  by Subject InstanceName; run;

    /* ---- Final consistency dataset ---- */
    data _rs_final;
        merge _rs_compare _cr_type _cr_classify;
        by Subject InstanceName;

        /* CR logic based on lesion type */
        if fb = 'Only lymph node'            and SLDIAM < 10 then result3 = 'CR';
        if fb = 'Only non lymph node'        and l1 ^= ''   then result3 = 'CR';
        if fb = 'Both non lymph node and lymph node'
                                             and l1 ^= '' and SLDIAM < 10 then result3 = 'CR';

        /* If PR and CR both triggered, prioritize PR (remove CR) */
        if result1 = 'PR' and result3 = 'CR' then do;
            q = '';
            result1 = '';
        end;

        /* CR inconsistency */
        if TRGRESP = 'CR' and result3 ^= 'CR' then q = 'Inconsistent';
        if TRGRESP ^= 'CR' and result3 = 'CR' then q = 'Inconsistent';

        label result3 = 'Calculated target response (CR)';

        if InstanceName = 'Screening' then TRGRESP = '';
        if TRGRESP = '-' then q = 'EDC Target Response is empty';
        if result = '' and result1 = '' and result2 = ''
           and result3 = '' and TRGRESP = 'NA' then q = '';
        if TRCDAT1 ^= '';
    run;

    /* Compress to final output */
    proc sql;
        create table _rs_output as
            select Subject, InstanceName, TRCDAT1, SLDIAM, minSLDIAM,
                   q4, cz, cz1, fb, l1,
                   result, result2, result1, result3,
                   TRGRESP, q
            from _rs_final;
    quit;

    /* ================================================================ */
    /*  PART B: Lesion-Level Method & Location Consistency               */
    /* ================================================================ */

    /* Import screening data with methods */
    data _con_scr;
        set _recist.&scrDS;
        rename &scrLocVar  = TULOC1
               &scrSiteVar = S_R1SIT
               &scrRecVar  = RecordPosition
               &scrMetVar  = TRMETHD1
               &scrMetSpecVar = TRMETHDO;
    run;
    data _con_fu;
        set _recist.&fuDS;
        rename &fuLocVar  = TULOC1
               &fuSiteVar = S_R1SIT
               &fuRecVar  = RecordPosition
               &fuMetVar  = TRMETHD1
               &fuMetSpecVar = TRMETHDO;
    run;

    /* Build coded location + method */
    data _con_scr;
        set _con_scr;
        TULOC9   = compress(cat(TULOC1, compress(S_R1SIT, , 'ak')), '09'x);
        TRMETHDF = compress(cat(TRMETHD1, compress(TRMETHDO, , 'ak')), '09'x);
        if TRMETHDO ^= '' then
            TRMETHDF = compress(cat(TRMETHD1, '(', compress(TRMETHDO), ')'));
        keep RecordPosition Subject InstanceName TULOC1 S_R1SIT TRMETHD1
             S_R1SIT TULOC9 TRMETHDF;
    run;

    data _con_fu;
        set _con_fu;
        TULOC9   = compress(cat(TULOC1, compress(S_R1SIT, , 'ak')), '09'x);
        TRMETHDF = compress(cat(TRMETHD1, compress(TRMETHDO, , 'ak')), '09'x);
        if TRMETHDO ^= '' then
            TRMETHDF = compress(cat(TRMETHD1, '(', compress(TRMETHDO), ')'));
        rename TRMETHD1 = TRMETHD3;
        keep RecordPosition Subject InstanceName TULOC1 S_R1SIT TRMETHD1
             S_R1SIT TULOC9 TRMETHDF;
    run;

    data _con_all;
        set _con_scr _con_fu;
        y2 = cats(Subject, TULOC9, S_R1SIT);
        y1 = cats(Subject, InstanceName);
        xq  = upcase(compress(TULOC1, , 'nk'));
        xq1 = upcase(compress(S_R1SIT, , 'nk'));
    run;

    /* Re-number records so screening and follow-up align */
    proc sort data=_con_all; by y1 xq xq1 RecordPosition; run;

    data _con_reindex;
        set _con_all;
        retain w;
        if first.y1 then w = 1;
        else w = w + 1;
        by y1;
        if RecordPosition = 0 then w = 0;
    run;

    data _con_all;
        set _con_reindex;
        RecordPositiona = RecordPosition;
        RecordPosition  = w;
    run;

    /* Count targets in screening vs follow-up */
    proc sort data=_con_all; by y2; run;

    data _con_pos;
        set _con_all;
        y1 = cats(Subject, InstanceName);
        keep Subject InstanceName RecordPosition y1 RecordPositiona;
    run;

    proc sort data=_con_pos; by y1 RecordPosition; run;

    data _con_max;
        set _con_pos;
        if last.y1;
        by y1;
        rename RecordPosition  = RecordPosition1;
        keep Subject InstanceName RecordPosition RecordPositiona;
        label RecordPosition = 'Target Number';
    run;

    data _con_scr_cnt;
        set _con_max;
        if InstanceName = "&scrLabel";
        rename RecordPosition1 = RecordPosition11;
        keep Subject RecordPosition1 RecordPositiona;
    run;

    proc sort data=_con_max; by Subject; run;
    proc sort data=_con_scr_cnt; by Subject; run;

    data _con_num_chk;
        merge _con_max _con_scr_cnt;
        by Subject;
        length q1 $200;
        if RecordPosition11 ^= RecordPosition1 and RecordPosition11 ^= '' then
            q1 = 'Amount Inconsistent';
        keep q1 RecordPosition11 Subject InstanceName;
    run;

    proc sort data=_con_num_chk; by Subject InstanceName; run;
    proc sort data=_con_all;      by Subject InstanceName; run;

    /* Method consistency check */
    data _con_all;
        merge _con_all _con_num_chk;
        by Subject InstanceName;
    run;

    proc sort data=_con_all; by y1; run;

    data _con_method;
        set _con_all;
        y2_check = lag(TRMETHDF);
        if first.y1 then y2_check = '';
        by y1;
    run;

    data _con_method;
        set _con_method;
        length q $200;
        if y2_check ^= TRMETHDF and y2_check ^= '' then q = 'Method Inconsistent';
        label q  = 'Method Result';
        label q1 = 'Number Result';
        drop y1 y2_check;
    run;

    /* Split screening and follow-up for cross-visit join */
    data _con_scr_dtl;
        set _con_method;
        if InstanceName = "&scrLabel";
        rename TRMETHDF = TRMETHD11;
        keep Subject TULOC1 S_R1SIT TRMETHDF TULOC9 RecordPosition
             RecordPosition11 RecordPositiona;
    run;

    data _con_fu_dtl;
        set _con_method;
        if InstanceName ^= "&scrLabel";
        TULOC11   = TULOC1;
        S_R1SIT1  = S_R1SIT;
        drop q TULOC1 S_R1SIT;
    run;

    data _con_fu_empty _con_fu_valid;
        set _con_fu_dtl;
        if TULOC9 = '' then output _con_fu_empty;
        else output _con_fu_valid;
    run;

    proc sort data=_con_scr_dtl;   by Subject RecordPosition; run;
    proc sort data=_con_fu_valid;  by Subject RecordPosition; run;
    proc sort data=_con_fu_empty;  by Subject; run;

    data _con_scr_dtl;
        merge _con_scr_dtl _con_fu_empty;
        by Subject;
    run;

    /* Cross-visit location match */
    data _con_location;
        length q $200;
        merge _con_scr_dtl _con_fu_valid;
        by Subject RecordPosition;
    run;

    /* Method consistency final labeling */
    data _con_location;
        set _con_location;
        if TRMETHD11 ^= TRMETHDF then q = 'Method Inconsistent';
        if TRMETHD11  = '' then q = 'Add new';
        if TRMETHDF    = '' then q = 'Treatment period Method is empty';
        if TULOC11 = '' and TULOC1 = '' then delete;
        drop TULOC9;
        label TRMETHDF = 'Method of Test or Examination';
        label RecordPosition    = 'Treatment period Record number';
        label RecordPosition11  = 'Screening sum of lesion';
        label TULOC1    = 'Screening Location';
        label TULOC11   = 'Treatment period Location';
    run;

    /* ---- Final consistency output ---- */
    proc sql;
        create table _con_output as
            select Subject, TULOC1, S_R1SIT, TRMETHD11, RecordPosition11,
                   InstanceName, RecordPosition, TULOC11, S_R1SIT1,
                   TRMETHDF, q1, q
            from _con_location;
    quit;

    /* ================================================================ */
    /*  Export to Excel                                                  */
    /* ================================================================ */

    %put NOTE: qcRecist — writing &_xlsxname..xlsx to &out;

    PROC EXPORT DATA=_rs_output
        OUTFILE = "&out\&_xlsxname..xlsx"
        DBMS = xlsx REPLACE label;
        SHEET = "Recist";
    RUN;

    PROC EXPORT DATA=_con_output
        OUTFILE = "&out\&_xlsxname..xlsx"
        DBMS = xlsx REPLACE label;
        SHEET = "Consistency";
    RUN;

    /* ---- Cleanup ---- */
    %_dropIfExists(dsn=_rs_scr);     %_dropIfExists(dsn=_rs_fu);
    %_dropIfExists(dsn=_rs_all);     %_dropIfExists(dsn=_rs_rank);
    %_dropIfExists(dsn=_scr_ref);    %_dropIfExists(dsn=_scr_base);
    %_dropIfExists(dsn=_rs_calc);    %_dropIfExists(dsn=_rs_clean);
    %_dropIfExists(dsn=_cr_chk);     %_dropIfExists(dsn=_cr_x);
    %_dropIfExists(dsn=_cr_disappear); %_dropIfExists(dsn=_cr_flag);
    %_dropIfExists(dsn=_edc_resp);   %_dropIfExists(dsn=_rs_compare);
    %_dropIfExists(dsn=_cr_merge);   %_dropIfExists(dsn=_cr_classify);
    %_dropIfExists(dsn=_cr_loc);     %_dropIfExists(dsn=_cr_ln);
    %_dropIfExists(dsn=_cr_nonln);   %_dropIfExists(dsn=_cr_both);
    %_dropIfExists(dsn=_cr_type);    %_dropIfExists(dsn=_rs_final);
    %_dropIfExists(dsn=_rs_output);  %_dropIfExists(dsn=_con_scr);
    %_dropIfExists(dsn=_con_fu);     %_dropIfExists(dsn=_con_all);
    %_dropIfExists(dsn=_con_reindex); %_dropIfExists(dsn=_con_pos);
    %_dropIfExists(dsn=_con_max);    %_dropIfExists(dsn=_con_scr_cnt);
    %_dropIfExists(dsn=_con_num_chk); %_dropIfExists(dsn=_con_method);
    %_dropIfExists(dsn=_con_scr_dtl); %_dropIfExists(dsn=_con_fu_dtl);
    %_dropIfExists(dsn=_con_fu_empty); %_dropIfExists(dsn=_con_fu_valid);
    %_dropIfExists(dsn=_con_location); %_dropIfExists(dsn=_con_output);

    libname _recist clear;

    %put NOTE: qcRecist — complete. Output: &out\&_xlsxname..xlsx;
%mend qcRecist;
