/****************************************************************************
 *  Test Script: %versionDiff
 *  Purpose: Create mock old/new library data and run versionDiff comparison.
 *
 *  Setup:
 *    1. Creates two temporary directories and assigns librefs:
 *       - oldlib  = old version (2 subjects, 2 datasets)
 *       - newlib  = new version (3 subjects, 2 datasets, with changes)
 *       - outlib  = output (temp)
 *    2. Runs %versionDiff to compare and export.
 *    3. (DDE colour test is skipped — Excel DDE requires interactive session.)
 *
 *  Expected changes detected:
 *    - DM: 1 New subject (USUBJID=003), 1 Update (USUBJID=001 data changed)
 *    - AE: 1 Inact record (only in old)
 ****************************************************************************/

/* ---- Clean up any pre-existing datasets ---- */
proc datasets library=work nolist kill; quit;

/* ---- Assign librefs ---- */
/* Using WORK sub-directories for portability; in real use these would be
   permanent directories. */

options dlcreatedir;
libname oldlib (work);
libname newlib (work);
libname outlib (work);

/* ---- Build OLD DM dataset ---- */
data oldlib.dm_old;
    infile datalines dsd;
    input USUBJID $ SITEID $ SEX $ AGE ARM $ RUNDATE :$8. _STATUE $;
datalines;
001,01,F,45,Active,20240101,Active
002,02,M,60,Placebo,20240101,Active
;

/* ---- Build NEW DM dataset ---- */
data newlib.dm_new;
    infile datalines dsd;
    input USUBJID $ SITEID $ SEX $ AGE ARM $ RUNDATE :$8. _STATUE $;
datalines;
001,01,F,45,Treatment,20240501,Active
002,02,M,60,Placebo,20240501,Active
003,03,F,52,Active,20240501,Active
;

/* ---- Build OLD AE dataset ---- */
data oldlib.ae_old;
    infile datalines dsd;
    input USUBJID $ AETERM $ AESEV $ AEOUT $ RUNDATE :$8. _STATUE $;
datalines;
001,Headache,MILD,RECOVERED,20240101,Active
002,Nausea,MODERATE,ONGOING,20240101,Active
002,Fatigue,MILD,RECOVERED,20240101,Inact
;

/* ---- Build NEW AE dataset ---- */
data newlib.ae_new;
    infile datalines dsd;
    input USUBJID $ AETERM $ AESEV $ AEOUT $ RUNDATE :$8. _STATUE $;
datalines;
001,Headache,MODERATE,RECOVERED,20240501,Active
002,Nausea,MODERATE,ONGOING,20240501,Active
003,Dizziness,MILD,ONGOING,20240501,Active
;

/* ---- Manually set libref metadata (workaround for WORK subdirs) ---- */
* Copy datasets to permanent-looking names so the macro can find them via dictionary.tables;
data newlib.DM;  set newlib.dm_new;  run;
data newlib.AE;  set newlib.ae_new;  run;
data oldlib.DM;  set oldlib.dm_old;  run;
data oldlib.AE;  set oldlib.ae_old;  run;

proc datasets library=newlib nolist;
    delete dm_new ae_new;
run;
proc datasets library=oldlib nolist;
    delete dm_old ae_old;
run;

/* ---- Run %versionDiff ---- */
title "Test: %versionDiff — DM + AE comparison";

%versionDiff(
    newLib  = newlib,
    oldLib  = oldlib,
    outLib  = outlib,
    keyvar  = USUBJID,
    exclVar = "_STATUE","RUNDATE",
    dropVar =
);

/* ---- Verify output datasets ---- */
title2 "Output library contents";
proc contents data=outlib._all_ nods; run;

title2 "DM — final version dataset";
proc print data=outlib.DM; run;

title2 "AE — final version dataset";
proc print data=outlib.AE; run;

title "Test complete.";
