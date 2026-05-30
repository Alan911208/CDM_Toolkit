/***************************************************************************/
/*  Test: csv2sas                                                           */
/*  Description: Creates a mock CSV file, imports it with both naming       */
/*               conventions (XLS2SAS and LABEL), and prints results.       */
/***************************************************************************/

/* Step 1: Include the macro */
%include "E:\Project\Tools_DM\sas_macros\11_csv2sas\csv2sas.sas";

/* Step 2: Create a mock CSV file using DATA _NULL_ with FILE statement   */
/*         Headers: "Subject ID", "AE Term$", "1st Dose Date"              */
/*         Three data rows with varying-length values                      */

%let _csvfile = %sysfunc(pathname(work))\test_ae_data.csv;

data _null_;
    file "&_csvfile" dlm=',' dsd lrecl=2000;
    /* Header row */
    put 'Subject ID,AE Term$,1st Dose Date';
    /* Data rows */
    put '101,Headache moderate,2023-01-15';
    put '102,Nausea$,2023-02-20';
    put '103,Dizziness (mild) and fatigue,2023-03-10';
run;

%put NOTE: Mock CSV written to &_csvfile;

/* Step 3: Import with naming=XLS2SAS (default) */
%csv2sas(file=&_csvfile, dsn=work.csv_xls2sas, naming=XLS2SAS);

title "csv2sas — naming=XLS2SAS";
proc print data=work.csv_xls2sas label noobs;
run;

title "csv2sas — naming=XLS2SAS — Contents";
proc contents data=work.csv_xls2sas;
run;
title;

/* Step 4: Import with naming=LABEL */
%csv2sas(file=&_csvfile, dsn=work.csv_label, naming=LABEL);

title "csv2sas — naming=LABEL";
proc print data=work.csv_label label noobs;
run;

title "csv2sas — naming=LABEL — Contents";
proc contents data=work.csv_label;
run;
title;

/* Step 5: Cleanup */
%_dropIfExists(dsn=work.csv_xls2sas);
%_dropIfExists(dsn=work.csv_label);
%_dropIfExists(dsn=work.__RST);
