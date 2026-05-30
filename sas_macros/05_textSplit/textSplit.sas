/*****************************************************************************
 *  Macro:  %textSplit
 *  Purpose: Split a long character variable (>200 chars) into N columns of
 *           max 200 characters each. Useful for CDISC text fields (e.g.,
 *           narrative text, comments) that exceed SAS character limits.
 *
 *  Parameters:
 *    dsn = Dataset name (required)
 *    var = Character variable to split (required)
 *
 *  Usage:
 *    %textSplit(dsn=work.ae, var=AETERM);
 *
 *  Internal helpers included: _dropIfExists
 *
 *  Source: DM_Toolbox.sas — Tool 5, Category 2
 *****************************************************************************/

/* ===== INTERNAL HELPER ===== */
%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;

/* ===== MAIN MACRO ===== */
%macro textSplit(dsn=, var=);
    %local _max _label;
    %if %length(&dsn) = 0 or %length(&var) = 0 %then %do;
        %put ERROR: textSplit — dsn and var are required;
        %return;
    %end;

    data _cut1;
        set &dsn;
        _n = ceilz(lengthn(ktrim(&var)) / 200);
    run;

    proc sql noprint;
        select max(_n), vlabel(&var) into :_max, :_label from _cut1;
    quit;

    %if &_max = . %then %let _max = 0;
    %if &_max = 0 %then %do;
        %put WARNING: textSplit — &var is empty in all rows; %return;
    %end;
    %put NOTE: textSplit — splitting &var into &_max segments;

    data &dsn(drop=_n _tot _rtot _rl _kl);
        set _cut1;
        _tot  = 1;
        _rtot = 1;
        array _aa $200 &var.1 - &var.&_max;
        %do _i = 1 %to &_max;
            if ktrim(substr(kleft(&var), _rtot, 200))
               = ktrim(substr(kleft(&var), _rtot, 199))
            then _rl = length(substr(kleft(&var), _rtot, 199));
            else _rl = length(substr(kleft(&var), _rtot, 200));
            _kl = klength(substr(kleft(&var), _rtot, _rl));
            if _kl > 0 then &var.&_i = strip(ksubstr(kleft(&var), _tot, _kl));
            _tot  = _tot  + _kl;
            _rtot = _rtot + _rl;
            label &var.&_i = "&&_label.&_i";
        %end;
        drop &var;
    run;

    %_dropIfExists(dsn=work._cut1);
%mend textSplit;
