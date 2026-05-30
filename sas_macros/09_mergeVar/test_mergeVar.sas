/*****************************************************************************
 * Test script for %mergeVar
 *
 * Creates 3 datasets in a test source directory:
 *   1. dm   — Subject-level info (usubjid, rficdtc, rfendtc)
 *   2. ae   — Adverse events (usubjid, aeseq, aeterm)
 *   3. lb   — Lab results (usubjid, lbseq, lbtestcd, lbstresc)
 *
 * Purpose: Merge rficdtc from dm into ae and lb by usubjid.
 *
 * Expected output:
 *   - dm: unchanged (same as input)
 *   - ae: augmented with rficdtc from dm
 *   - lb: augmented with rficdtc from dm
 *****************************************************************************/

/* Include the macro */
%include "E:/Project/Tools_DM/sas_macros/09_mergeVar/mergeVar.sas";

title "===== Test: %mergeVar =====";

/* Step 1: Set up test directories */
%let srcPath = %sysfunc(getoption(WORK))/mergeVar_src;
%let tgtPath = %sysfunc(getoption(WORK))/mergeVar_tgt;

options noxwait;
x "rmdir /s /q &srcPath 2>NUL";
x "rmdir /s /q &tgtPath 2>NUL";
x "mkdir &srcPath";
x "mkdir &tgtPath";

libname _src "&srcPath";
libname _tgt "&tgtPath";

/* Create DM dataset — subject demographics */
data _src.dm;
    input USUBJID $ RFICDTC $ RFENDTC $;
    datalines;
SUBJ001 2023-01-15 2023-06-30
SUBJ002 2023-02-20 2023-07-15
SUBJ003 2023-03-10 2023-08-20
;
run;

/* Create AE dataset — adverse events */
data _src.ae;
    input USUBJID $ AESEQ AETERM $;
    datalines;
SUBJ001 1 Headache
SUBJ001 2 Nausea
SUBJ002 1 Fatigue
SUBJ003 1 Dizziness
;
run;

/* Create LB dataset — lab results */
data _src.lb;
    input USUBJID $ LBSEQ LBTESTCD $ LBSTRESC $;
    datalines;
SUBJ001 1 ALT 25
SUBJ001 2 AST 30
SUBJ002 1 ALT 22
SUBJ003 1 ALT 18
SUBJ003 2 AST 20
;
run;

/* Log input datasets */
title "Input: dm";
proc print data=_src.dm noobs; run;

title "Input: ae (NO rficdtc)";
proc print data=_src.ae noobs; run;

title "Input: lb (NO rficdtc)";
proc print data=_src.lb noobs; run;

/* Step 2: Run mergeVar — merge rficdtc from dm into all datasets */
%mergeVar(form=dm, vars=rficdtc, key=usubjid, src=&srcPath, tgt=&tgtPath);

/* Step 3: Print output datasets */
title "Output: dm (unchanged, still has rficdtc and rfendtc)";
proc print data=_tgt.dm noobs; run;

title "Output: ae (now has rficdtc from dm)";
proc print data=_tgt.ae noobs; run;

proc contents data=_tgt.ae;
run;

title "Output: lb (now has rficdtc from dm)";
proc print data=_tgt.lb noobs; run;

proc contents data=_tgt.lb;
run;

/* Step 4: Verify rficdtc values match across datasets */
title "Verification: Check rficdtc consistency across ae and lb";
proc sql;
    select a.usubjid, a.rficdtc as ae_rficdtc, l.rficdtc as lb_rficdtc
    from _tgt.ae a
    inner join _tgt.lb l on a.usubjid = l.usubjid
    group by a.usubjid;
quit;

/* Step 5: Cleanup */
libname _src clear;
libname _tgt clear;

title "===== End Test: %mergeVar =====";
