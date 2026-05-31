/****************************************************************************
 *  Test Script: %qcRecist
 *  Purpose: Create mock RECIST 1.1 screening, follow-up, and response
 *           datasets and run qcRecist to verify calculations.
 *
 *  Mock data scenario (2 subjects):
 *    Subject 001 (Baseline SLD=60mm, 2 target lesions):
 *      - Follow-up Visit 1: SLD=45mm (25% decrease) → calculated PR
 *        EDC says PR → consistent (no flag)
 *      - Follow-up Visit 2: SLD=90mm (50% increase, +30mm) → calculated PD
 *        EDC says SD → inconsistent (FLAG)
 *    Subject 002 (Baseline SLD=100mm, 3 target lesions):
 *      - Follow-up Visit 1: SLD=95mm (5% decrease, within SD range) → calculated SD
 *        EDC says SD → consistent
 *      - Follow-up Visit 2: SLD=30mm (70% decrease) → calculated PR
 *        EDC says PR → consistent
 *
 *  Note: The real RECIST calculation is more nuanced — this test confirms
 *        the macro runs and handles basic scenarios correctly.
 ****************************************************************************/

/* ---- Clean workspace ---- */
proc datasets library=work nolist kill; quit;

/* ---- Set up output path ---- */
%let _outPath = %sysfunc(pathname(work));

/* ---- Build SCREENING lesion dataset ---- */
data work._1recist1;
    length Subject $10 InstanceName $20 TRCDAT1 $9;
    input Subject $ InstanceName $ TRCDAT1 $ LDIAM SLDIAM
          TULOC1 $ 1-20 S_R1SIT $ 21-40 RecordPosition TRMETHD1 $ TRMETHDO $;
datalines;
001       Screening   01Jan2024  30   60   Right Lung Upper     Segment 1          1  CT         Biopsy
001       Screening   01Jan2024  30   60   Left Lung Lower      Segment 3          2  CT
002       Screening   15Jan2024  40  100   Liver Segment 4      Subcapsular        1  MRI
002       Screening   15Jan2024  30  100   Liver Segment 7      Central            2  MRI
002       Screening   15Jan2024  30  100   Lung Right Lower     Posterior           3  MRI
;

/* ---- Build FOLLOW-UP lesion dataset ---- */
data work._3recist1;
    length Subject $10 InstanceName $20 TRCDAT1 $9;
    input Subject $ InstanceName $ TRCDAT1 $ LDIAM SLDIAM
          TULOC1 $ 1-20 S_R1SIT $ 21-40 RecordPosition TRMETHD1 $ TRMETHDO $;
datalines;
001       Week 8       01Mar2024  22   45   Right Lung Upper     Segment 1          1  CT         Biopsy
001       Week 8       01Mar2024  23   45   Left Lung Lower      Segment 3          2  CT
001       Week 16      01May2024  45   90   Right Lung Upper     Segment 1          1  CT         Biopsy
001       Week 16      01May2024  45   90   Left Lung Lower      Segment 3          2  CT
002       Week 8       01Mar2024  38   95   Liver Segment 4      Subcapsular        1  MRI
002       Week 8       01Mar2024  28   95   Liver Segment 7      Central            2  MRI
002       Week 8       01Mar2024  29   95   Lung Right Lower     Posterior           3  MRI
002       Week 16      01May2024  12   30   Liver Segment 4      Subcapsular        1  MRI
002       Week 16      01May2024  10   30   Liver Segment 7      Central            2  MRI
002       Week 16      01May2024   8   30   Lung Right Lower     Posterior           3  MRI
;

/* ---- Build EDC RESPONSE dataset ---- */
data work.overall;
    length Subject $10 InstanceName $20 TRGRESP $4;
    input Subject $ InstanceName $ TRGRESP $;
datalines;
001       Week 8       PR
001       Week 16      SD
002       Week 8       SD
002       Week 16      PR
;

/* ---- Run %qcRecist ---- */
title "Test: %qcRecist — RECIST 1.1 Response Consistency";

%qcRecist(
    lib      = &_outPath,
    out      = &_outPath,
    scrDS    = _1recist1,
    fuDS     = _3recist1,
    respDS   = overall,
    respVar  = TRGRESP,
    scrDateVar = TRCDAT1,
    fuDateVar  = TRCDAT1,
    scrLongVar = LDIAM,
    fuLongVar  = LDIAM,
    scrSumVar  = SLDIAM,
    fuSumVar   = SLDIAM,
    scrLocVar  = TULOC1,
    fuLocVar   = TULOC1,
    scrSiteVar = S_R1SIT,
    fuSiteVar  = S_R1SIT,
    scrRecVar  = RecordPosition,
    fuRecVar   = RecordPosition,
    scrMetVar  = TRMETHD1,
    fuMetVar   = TRMETHD1,
    scrMetSpecVar = TRMETHDO,
    fuMetSpecVar  = TRMETHDO,
    scrLabel   = Screening
);

/* ---- Verify results ---- */

title2 "Recist Output (target response comparison)";
%if %sysfunc(exist(_rs_output, data)) %then %do;
    proc print data=_rs_output noobs;
        var Subject InstanceName SLDIAM minSLDIAM cz1 result1 result2 result TRGRESP q;
    run;
%end;

title2 "Consistency Output (lesion method/location)";
%if %sysfunc(exist(_con_output, data)) %then %do;
    proc print data=_con_output noobs;
        var Subject TULOC1 TULOC11 InstanceName q1 q;
    run;
%end;

title "Test complete.";
