/***************************************************************************/
/*  Macro: sas2xls                                                          */
/*  Purpose: Export a SAS dataset to a specific Excel sheet using DDE       */
/*           (Dynamic Data Exchange). This is a LEGACY approach that        */
/*           requires Microsoft Excel running on a Windows SAS session.     */
/*                                                                          */
/*           For modern SAS (9.4+), prefer sas2xlsx (macro 13) which uses   */
/*           PROC EXPORT with dbms=xlsx and does not require Excel.         */
/*                                                                          */
/*  Prerequisites:                                                          */
/*    - Windows operating system                                            */
/*    - Microsoft Excel installed and running at time of execution          */
/*    - DDE must be available in the SAS session                            */
/*                                                                          */
/*  Parameters:                                                              */
/*    inset    = Input SAS dataset (one-level or two-level name)            */
/*               Use one-level name for WORK library (default: last         */
/*               created dataset)                                           */
/*    outsheet = Target Excel sheet name (default: Sheet1)                  */
/*    outfile  = Target Excel workbook path (required)                      */
/*               If the workbook does not exist, DDE will attempt to        */
/*               create it via an Excel system command.                     */
/*                                                                          */
/*  Usage example:                                                           */
/*    %sas2xls(inset=work.AE, outsheet=AE_Data,                              */
/*             outfile=C:\output\ae_report.xls);                             */
/*                                                                          */
/*  Notes:                                                                   */
/*    - Not all SAS installations support DDE (Server/Cite environments)    */
/*    - Excel must be opened before calling this macro                      */
/*    - The DDE triplet format is: Excel|System                               */
/*    - This macro does NOT handle advanced formatting (fonts, colors, etc.) */
/*                                                                          */
/*  Author: DM Tools                                                        */
/*  Version: 1.0                                                            */
/***************************************************************************/

%macro sas2xls(inset=, outsheet=Sheet1, outfile=);

    /* Validate parameters */
    %if %length(&inset) = 0 %then %do;
        %put ERROR: sas2xls - inset parameter is required;
        %return;
    %end;

    %if %length(&outfile) = 0 %then %do;
        %put ERROR: sas2xls - outfile parameter is required;
        %return;
    %end;

    /* Resolve dataset reference */
    %local _ds _lib _mem _libpath;
    %if %index(&inset, .) > 0 %then %do;
        %let _lib = %scan(&inset, 1, .);
        %let _mem = %scan(&inset, 2, .);
    %end;
    %else %do;
        %let _lib = work;
        %let _mem = &inset;
    %end;

    /* Verify the dataset exists */
    %if not %sysfunc(exist(&_lib..&_mem, data)) %then %do;
        %put ERROR: sas2xls - Dataset &_lib..&_mem does not exist;
        %return;
    %end;

    %put NOTE: sas2xls - Exporting &_lib..&_mem via DDE to Excel;
    %put NOTE: sas2xls - Target: &outfile, Sheet: &outsheet;

    /*
     * The standard DDE approach is:
     *
     * 1. Open Excel workbook via DDE system command:
     *    filename _ddecmd dde 'Excel|System';
     *    data _null_;
     *        file _ddecmd;
     *        put '[OPEN("' "&outfile" '")]';
     *    run;
     *
     * 2. Write data to the target sheet:
     *    filename _ddedat dde "Excel|&outsheet!R1C1:R9999C99";
     *    data _null_;
     *        set &_lib..&_mem;
     *        file _ddedat;
     *        put var1 var2 ... ;
     *    run;
     *
     * 3. Save and close:
     *    data _null_;
     *        file _ddecmd;
     *        put '[SAVE()]';
     *        put '[CLOSE()]';
     *    run;
     *
     * Because DDE requires an Excel session to be already running
     * (often started via X command) and the exact DDE triplets can
     * vary by SAS/Excel version, this macro provides a commented
     * template below. Uncomment and adjust for your environment.
     */

    %put WARNING: sas2xls — DDE approach requires Excel running on Windows;
    %put WARNING: sas2xls — This macro is a template; review DDE triplets for your SAS/Excel version;
    %put NOTE: sas2xls - ============================================;
    %put NOTE: sas2xls - DDE export template (uncomment to use):;
    %put NOTE: sas2xls - ============================================;

    /*
    *** Begin DDE export block — uncomment and adjust as needed ***

    * --- Open workbook --- *;
    filename excelcmd dde 'Excel|System';
    data _null_;
        file excelcmd;
        put '[OPEN("' "&outfile" '")]';
    run;

    * --- Write variable names as header row --- *;
    filename excelsht dde "Excel|&outsheet!R1C1:R1C999";
    data _null_;
        file excelsht;
        put 'SUBJID' '09'x 'AGE' '09'x 'SEX' '09'x 'RACE';
    run;

    * --- Write data rows --- *;
    filename excelsht dde "Excel|&outsheet!R2C1:R9999C99";
    data _null_;
        set &_lib..&_mem;
        file excelsht;
        put SUBJID '09'x AGE '09'x SEX $ '09'x RACE $;
    run;

    * --- Save and close --- *;
    data _null_;
        file excelcmd;
        put '[SAVE()]';
        put '[CLOSE()]';
    run;

    *** End DDE export block ***
    */

    %put NOTE: sas2xls - Template complete. Review and customize the DDE code block above.;
%mend sas2xls;
