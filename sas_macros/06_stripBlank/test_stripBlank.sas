/*****************************************************************************
 * Test script for %stripBlank
 *
 * Creates a CDM-style dataset with:
 *   - line: page line number
 *   - Status: record status (1=active, 0=blank, .=missing)
 *   - pStatus: page-level status ('Added    ', 'Modified ', 'Deleted  ')
 *   - SEQCD: numeric sequence code (all missing in this test)
 *   - SEQTEXT: text sequence (all missing in this test)
 *   - Data columns: USUBJID, VISIT, VISITNUM, LBTESTCD, LBSTRESC
 *
 * The raw dataset contains:
 *   - 2 blank placeholder rows (line=., all data=. / '')
 *   - 1 "Added" row that should be removed (pStatus='Added    ', Status=0)
 *   - 4 valid data rows
 *
 * Expected output:
 *   - blank rows removed
 *   - "Added" placeholder row removed
 *   - SEQCD column dropped (all missing)
 *   - SEQTEXT column dropped (all missing)
 *   - pStatus column dropped (keepPS=NO by default)
 *   - SEQNUM column dropped (all missing, numeric format)
 *   - 4 output rows
 *****************************************************************************/

/* Include the macro */
%include "E:/Project/Tools_DM/sas_macros/06_stripBlank/stripBlank.sas";

title "===== Test: %stripBlank =====";

/* Step 1: Create a CDM-style raw dataset with blank/placeholder rows */
data work.raw_cdm;
    /* Row 1: valid data */
    line=1; Status=1; pStatus='Modified '; SEQNUM=.; SEQCD=.; SEQTEXT='';
    USUBJID='SUBJ-001'; VISIT='Week 0'; VISITNUM=0; LBTESTCD='ALT'; LBSTRESC='25'; output;
    /* Row 2: valid data */
    line=2; Status=1; pStatus='          '; SEQNUM=.; SEQCD=.; SEQTEXT='';
    USUBJID='SUBJ-001'; VISIT='Week 0'; VISITNUM=0; LBTESTCD='AST'; LBSTRESC='30'; output;
    /* Row 3: blank placeholder row — should be removed */
    line=.; Status=.; pStatus='          '; SEQNUM=.; SEQCD=.; SEQTEXT='';
    USUBJID=''; VISIT=''; VISITNUM=.; LBTESTCD=''; LBSTRESC=''; output;
    /* Row 4: valid data */
    line=3; Status=1; pStatus='          '; SEQNUM=.; SEQCD=.; SEQTEXT='';
    USUBJID='SUBJ-001'; VISIT='Week 2'; VISITNUM=2; LBTESTCD='ALT'; LBSTRESC='22'; output;
    /* Row 5: Added placeholder — has Status=0 and pStatus='Added' */
    line=4; Status=0; pStatus='Added    '; SEQNUM=.; SEQCD=.; SEQTEXT='';
    USUBJID='SUBJ-001'; VISIT='Week 2'; VISITNUM=2; LBTESTCD='BILI'; LBSTRESC=''; output;
    /* Row 6: valid data */
    line=5; Status=1; pStatus='          '; SEQNUM=.; SEQCD=.; SEQTEXT='';
    USUBJID='SUBJ-002'; VISIT='Week 0'; VISITNUM=0; LBTESTCD='ALT'; LBSTRESC='18'; output;
    /* Row 7: blank placeholder row — should be removed */
    line=.; Status=.; pStatus='          '; SEQNUM=.; SEQCD=.; SEQTEXT='';
    USUBJID=''; VISIT=''; VISITNUM=.; LBTESTCD=''; LBSTRESC=''; output;
run;

/* Log the input */
title "Input: raw_cdm (7 rows, including blank/placeholder rows)";
proc print data=work.raw_cdm noobs;
run;

/* Step 2: Run stripBlank */
%stripBlank(inlib=work, outlib=work, insets=raw_cdm, outsets=clean_cdm, idxVar=SEQNUM, keepPS=NO);

/* Step 3: Print the cleaned result */
title "Output: clean_cdm (expected 4 rows, no SEQNUM/SEQCD/SEQTEXT/pStatus)";
proc print data=work.clean_cdm noobs;
run;

/* Step 4: Check contents */
proc contents data=work.clean_cdm;
run;

/* Step 5: Verify row count */
proc sql;
    select count(*) as N_rows from work.clean_cdm;
quit;

title "===== End Test: %stripBlank =====";
