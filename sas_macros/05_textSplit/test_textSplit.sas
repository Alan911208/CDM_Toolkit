/*****************************************************************************
 *  Test Script: %textSplit
 *  Creates a dataset with a 500-character text variable, runs textSplit,
 *  and verifies that 3 output columns are created with correct content.
 *****************************************************************************/

/* Include the macro definition (includes internal helper _dropIfExists) */
%include "E:\Project\Tools_DM\sas_macros\05_textSplit\textSplit.sas";

/* Step 1: Create mock dataset with a long text variable.
   We build a 500-character string by repeating a pattern.
   Each repetition of "ABCDEFGHIJ" is 10 chars, 50 reps = 500 chars */
data work.longtext;
    length ID 8 TEXT $500;
    /* Row 1: exactly 500 characters */
    TEXT = repeat("ABCDEFGHIJ", 49);  /* 50 * 10 = 500 chars */
    ID = 1;
    output;

    /* Row 2: 450 characters (should still split into 3 segments,
       but the 3rd segment will be shorter) */
    TEXT = repeat("0123456789", 44);  /* 45 * 10 = 450 chars */
    ID = 2;
    output;

    /* Row 3: empty text (should have _n = 0 but not cause errors,
       macro uses max(_n) so it still produces the needed columns) */
    TEXT = "";
    ID = 3;
    output;
run;

/* Step 2: Verify input — print the first 50 characters */
%put === Input: each row's TEXT length ===;
data _null_;
    set work.longtext;
    put "ID=" ID "LENGTH=" lengthn(TEXT);
run;

/* Step 3: Run textSplit */
%put === Running textSplit on TEXT variable ===;
%textSplit(dsn=work.longtext, var=TEXT);

/* Step 4: Check results */
%put === After textSplit: verify structure ===;
proc contents data=work.longtext varnum; run;

%put === After textSplit: verify content (first 30 chars of each segment) ===;
data _null_;
    set work.longtext;
    put "--- ID=" ID " ---";
    put "TEXT1 (first 30)=" substr(TEXT1,1,30);
    if ID <= 2 then put "TEXT2 (first 30)=" substr(TEXT2,1,30);
    if ID <= 2 then put "TEXT3 (first 30)=" substr(TEXT3,1,30);
run;

/* Step 5: Verify segment lengths */
proc sql;
    select ID,
           lengthn(TEXT1) as len1,
           lengthn(TEXT2) as len2,
           lengthn(TEXT3) as len3
    from work.longtext;
quit;

/* Step 6: Verify concatenation equals original */
data _null_;
    set work.longtext;
    combined = cats(of TEXT1-TEXT3);
    put "ID=" ID " Combined length=" lengthn(combined);
    /* For IDs 1 and 2, combined should match original 500/450 chars */
run;

/* Cleanup */
proc datasets lib=work nolist;
    delete longtext;
quit;
