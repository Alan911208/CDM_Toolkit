/*****************************************************************************
 *  Test Script: %varExist
 *  Tests the varExist macro by checking variables in a mock dataset.
 *  Since this cannot be executed without SAS, this script is structured to
 *  be clearly correct and well-commented for manual review.
 *****************************************************************************/

/* Include the macro definition */
%include "E:\Project\Tools_DM\sas_macros\01_varExist\varExist.sas";

/* Step 1: Create a mock dataset with variables A, B, C */
data work.demo;
    A = 1; B = 2; C = 3;
    output;
    A = 4; B = 5; C = 6;
    output;
run;

/* Step 2: Test existence of variable A — expected &rc = 1 */
%put === Test 1: Check variable A (should exist) ===;
%varExist(dsn=work.demo, var=A);
%put NOTE: varExist for A returned rc=&rc (expected: 1);

/* Step 3: Test existence of variable B — expected &rc = 1 */
%put === Test 2: Check variable B (should exist) ===;
%varExist(dsn=work.demo, var=B);
%put NOTE: varExist for B returned rc=&rc (expected: 1);

/* Step 4: Test existence of variable C — expected &rc = 1 */
%put === Test 3: Check variable C (should exist) ===;
%varExist(dsn=work.demo, var=C);
%put NOTE: varExist for C returned rc=&rc (expected: 1);

/* Step 5: Test existence of variable D (not in dataset) — expected &rc = 0 */
%put === Test 4: Check variable D (should NOT exist) ===;
%varExist(dsn=work.demo, var=D);
%put NOTE: varExist for D returned rc=&rc (expected: 0);

/* Step 6: Test with a non-existent dataset — expected WARNING */
%put === Test 5: Check non-existent dataset (should produce WARNING) ===;
%varExist(dsn=work.noSuchDS, var=X);
%put NOTE: varExist for non-existent ds returned rc=&rc (expected: 0);

/* Step 7: Test with missing parameters — expected ERROR */
%put === Test 6: Missing dsn parameter (should produce ERROR) ===;
%varExist(dsn=, var=X);

%put === Test 7: Missing var parameter (should produce ERROR) ===;
%varExist(dsn=work.demo, var=);

/* Step 8: Conditional usage demonstration */
%put === Test 8: Conditional usage ===;
%varExist(dsn=work.demo, var=A);
%if &rc = 1 %then %put NOTE: Variable A detected — conditional branch works;
%else %put ERROR: Variable A not detected — unexpected;

%varExist(dsn=work.demo, var=Z);
%if &rc = 0 %then %put NOTE: Variable Z not detected — conditional branch works;
%else %put ERROR: Variable Z detected — unexpected;

/* Cleanup */
proc datasets lib=work nolist; delete demo; quit;
