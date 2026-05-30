/*****************************************************************************
 *  Macro:  %varExist
 *  Purpose: Test whether a variable exists in a dataset.
 *           Returns &rc = 0 (not found) or 1 (found) as a GLOBAL macro var.
 *
 *  Parameters:
 *    dsn = Dataset name (required)
 *    var = Variable name (required)
 *
 *  Usage:
 *    %varExist(dsn=work.demo, var=A);
 *    %if &rc = 1 %then %do; ... %end;
 *
 *  Source: DM_Toolbox.sas — Tool 1, Category 1
 *****************************************************************************/

%macro varExist(dsn=, var=);
    %global rc;
    %local _ds _vn _fid _pos;
    %let rc = 0;
    %if %length(&dsn) = 0 or %length(&var) = 0 %then %do;
        %put ERROR: varExist — dsn and var are required;
        %return;
    %end;
    %let _fid = %sysfunc(open(&dsn,i));
    %if &_fid > 0 %then %do;
        %let _pos = %sysfunc(varnum(&_fid,&var));
        %if &_pos > 0 %then %let rc = 1;
        %let _fid = %sysfunc(close(&_fid));
    %end;
    %else %put WARNING: varExist — cannot open dataset &dsn;
%mend varExist;
