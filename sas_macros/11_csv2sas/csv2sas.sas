/***************************************************************************/
/*  Macro: csv2sas                                                         */
/*  Purpose: Import CSV file into SAS dataset with intelligent variable    */
/*           naming. 3-pass import: (1) parse headers + sanitise names,    */
/*           (2) detect max column widths, (3) read data with optimal      */
/*           lengths.                                                       */
/*  Parameters:                                                             */
/*    file   = Full path to the CSV file (required)                        */
/*    dsn    = Output dataset name (default: csv2sas)                      */
/*    naming = Variable naming convention: XLS2SAS (default) or LABEL      */
/*             XLS2SAS: _c1_, _c2_, ... with labels from header            */
/*             LABEL  : sanitized header names with labels from header     */
/*    lrecl  = Logical record length for INFILE (default: 5000)            */
/*  Author: DM Tools                                                       */
/*  Version: 1.0                                                           */
/***************************************************************************/

%macro _dropIfExists(dsn=);
    %if %sysfunc(exist(&dsn,data)) %then %do;
        proc sql noprint; drop table &dsn; quit;
    %end;
%mend _dropIfExists;

%macro csv2sas(file=, dsn=csv2sas, naming=XLS2SAS, lrecl=5000);
    %local _nvars _i;

    %if %length(&file) = 0 %then %do;
        %put ERROR: csv2sas - file parameter is required;
        %return;
    %end;

    options nomprint notes nosource2 nosource;

    /*** Pass 1: Read header row, sanitise names, store in __RST ***/
    data __RST(drop=_i _ascii);
        length _lbl _name $200;
        infile "&file" firstobs=1 obs=1 dsd lrecl=&lrecl;
        input _lbl $ @@ ;
        _err = 0;
        do _i = 1 to length(trim(_lbl));
            _ascii = rank(substr(trim(_lbl), _i, 1));
            if _i = 1 and (48 <= _ascii <= 57) then _ascii = 95;
            else if _ascii = 32 then _ascii = 95;
            if (48 <= _ascii <= 57) or _ascii = 95
               or (65 <= _ascii <= 90) or (97 <= _ascii <= 122)
            then do;
                if _i = 1 then _name = byte(_ascii);
                else _name = trim(_name) || byte(_ascii);
            end;
            else _err = 1;
        end;
        _dde = compress("_c" || _N_ || "_");
        if _name = '' or _err = 1 then _name = compress("_col" || _N_ - 1);
        call symputx("_nvars", _N_);
    run;

    /*** Pass 2: Scan data rows to detect maximum column widths ***/
    data _null_;
        retain %do _i = 1 %to &_nvars; _len&_i 0 %end; ;
        length %do _i = 1 %to &_nvars; _C&_i $255 %end; ;
        infile "&file" firstobs=2 dsd lrecl=&lrecl end=_stop;
        input %do _i = 1 %to &_nvars; _C&_i %end; ;
        %do _i = 1 %to &_nvars;
            if length(strip(_C&_i)) > _len&_i then _len&_i = length(strip(_C&_i));
        %end;
        if _stop then do;
            %do _i = 1 %to &_nvars;
                call symputx("_len&_i", _len&_i);
            %end;
        end;
    run;

    /*** Pass 3: Read data with optimal column lengths ***/
    data &dsn;
        length %do _i = 1 %to &_nvars; _C&_i $&&_len&_i %end; ;
        infile "&file" firstobs=2 dsd lrecl=&lrecl;
        input %do _i = 1 %to &_nvars; _C&_i %end; ;
    run;

    /*** Apply rename and label using CALL EXECUTE ***/
    data _null_;
        set __RST end=_stop;
        if _N_ = 1 then call execute("proc datasets lib=work nolist nodetails nowarn nofs; modify &dsn;");
        %if %upcase(&naming) = XLS2SAS %then %do;
            call execute("rename _c" || strip(_N_) || "= " || strip(_dde) || ";");
            call execute("label  "   || strip(_dde) || "= " || quote(strip(_lbl)) || ";");
        %end; %else %do;
            call execute("rename _c" || strip(_N_) || "= " || strip(_name) || ";");
            call execute("label  "   || strip(_name) || "= " || quote(strip(_lbl)) || ";");
        %end;
        if _stop then call execute("quit;");
    run;

    /*** Cleanup ***/
    %_dropIfExists(dsn=work.__RST);
    options source2 source;

    %put NOTE: csv2sas - Imported %sysfunc(putn(&_nvars, best.)) variables into dataset &dsn;
%mend csv2sas;
