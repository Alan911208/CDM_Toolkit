/*****************************************************************************
 * Test Script: runDateDiff
 * Purpose: Verify that %runDateDiff correctly distinguishes changed rows
 *          from unchanged rows and assigns rundate accordingly.
 *
 * Test Scenario:
 *   - OLD dataset (dm_old): 3 subjects with baseline data
 *   - NEW dataset (dm_new): same 3 subjects, but subject 001-0002 has
 *     a changed ARMCD, and 001-0004 is completely new
 *   - Expected: unchanged rows keep the old rundate ('01JAN2020'd);
 *     changed/new rows get today's date (from &sysdate9.)
 *****************************************************************************/

* ---- Step 1: Include the macro ----;
%include "&_projectRoot\sas_macros\17_runDateDiff\runDateDiff.sas";

* ---- Step 2: Create mock old dataset ----;
data dm_old;
    length usubjid $10 siteid $4 armcd $8;
    usubjid = "001-0001"; siteid = "S001"; armcd = "TRT"; rundate = '01JAN2020'd; output;
    usubjid = "001-0002"; siteid = "S001"; armcd = "PBO"; rundate = '01JAN2020'd; output;
    usubjid = "001-0003"; siteid = "S002"; armcd = "TRT"; rundate = '01JAN2020'd; output;
    format rundate date9.;
run;

* ---- Step 3: Create mock new dataset (some changes) ----;
data dm_new;
    length usubjid $10 siteid $4 armcd $8;
    * Row 1: unchanged;
    usubjid = "001-0001"; siteid = "S001"; armcd = "TRT"; output;
    * Row 2: ARMCD changed from PBO to TRT;
    usubjid = "001-0002"; siteid = "S001"; armcd = "TRT"; output;
    * Row 3: unchanged;
    usubjid = "001-0003"; siteid = "S002"; armcd = "TRT"; output;
    * Row 4: completely new subject;
    usubjid = "001-0004"; siteid = "S003"; armcd = "PBO"; output;
run;

* ---- Step 4: Run runDateDiff ----;
%runDateDiff(oldlib=work, oldset=dm_old,
             newlib=work, newset=dm_new,
             outlib=work, outset=dm_result,
             keyvar=usubjid,
             exvar=rundate);

* ---- Step 5: Verify results ----;

title "Test 1: Output dataset — all 4 rows present";
proc print data=dm_result noobs;
    format rundate date9.;
run;

title "Test 2: Unchanged rows (001-0001, 001-0003) — should keep 01JAN2020";
proc print data=dm_result noobs;
    where usubjid in ("001-0001", "001-0003");
    format rundate date9.;
run;

title "Test 3: Changed row (001-0002) — should have TODAY's date, not 01JAN2020";
proc print data=dm_result noobs;
    where usubjid = "001-0002";
    format rundate date9.;
run;

title "Test 4: New row (001-0004) — should have TODAY's date";
proc print data=dm_result noobs;
    where usubjid = "001-0004";
    format rundate date9.;
run;

title "Test 5: Summary — check rundate distribution";
proc freq data=dm_result;
    tables rundate;
    format rundate date9.;
run;

* ---- Step 6: Clean up ----;
proc datasets lib=work nolist;
    delete dm_old dm_new dm_result;
quit;
title;
