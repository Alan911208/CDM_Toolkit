/**************************************************************
* Program:      CSV2SAS
* Function:     Reading CSV file into SAS datset 
* Result:       SAS Dataset
* Date:         Feb 24, 2018
* Programmer:   Fangqiu Xiao
*
*
* To Replace %xls2SAS or proc import;
* Clinflash Applications for Clinical Data Management
*
* Sample Usage:

  %CSV2SAS(fCSV=D:\Query.csv,outset=query);

  %CSV2SAS(fCSV=D:\Query.csv,outset=query,type=);

*
* Parameters:
* fCSV=physical CSV file name 
* Outset=sas dataset name
* type=XLS2SAS -> similiar to %XLS2SAS, otherwise  PROC IMPORT;

* See also codes to Output dataset to CSV file:

ods listing close; 
ods csv file='D:\Query.csv'; 
	proc print data=query noobs label; run; 
ods csv close;
ods listing; 


* All Right Reserved 
*
* Change history:
* May 31, 2018 add parameter LRECL for large size csv files
* 
*
***************************************************************/

%macro CSV2SAS(fCSV=,outset=CSV2SAS,type=XLS2SAS,LRECL=5000);
	options nomprint notes nosource2 nosource;

	data __RST(drop=i ascii myERR);
		length lstLabel lstName $200;
		infile "&fCSV" firstobs=1 obs=1 dsd LRECL=&LRECL;
		input lstLabel $ @@ ;
		myERR=0;
		do i=1 to length(trim(lstLabel));
			ascii=rank(substr(trim(lstLabel),i,1));
			if i=1 then do;
				if (ascii ge 48 and ascii le 57) then ascii=95;
			end; else if ascii=32 then ascii=95;
			if (ascii ge 48 and ascii le 57) or ascii=95 or (ascii ge 65 and ascii le 90) or (ascii ge 97 and ascii le 122)  then do;
				if i=1 then lstName =byte(ascii);else lstName =trim(lstName)||byte(ascii);
			end; else myErr=1;
		end;
		ddeName=compress("_c"||_N_||"_");
		if lstName='' or myERR=1 then lstName=compress("_col"||_N_-1);
		call symput("nVars",compress(_N_));
	run;

	data _null_;
		retain _len1-_len&nVars 0;
		length %do i=1 %to &nVars; _C&i %end; $255;
		infile "&fCSV" firstobs=2 dsd LRECL=&LRECL end=stop ;
	    input %do i=1 %to &nVars; _C&i %end;;
		%do i=1 %to &nVars;if length(strip(_C&i)) > _len&i then   _len&i=length(strip(_C&i));%end;
		if stop then do;
			%do i=1 %to &nVars; call symput("_len&i", compress(_len&i)); %end;
		end;
	run;

	data &outset;
		length %do i=1 %to &nVars; _C&i $ &&_len&i %end;;
		infile "&fCSV" firstobs=2  dsd LRECL=&LRECL end=stop;
	    input %do i=1 %to &nVars; _C&i %end;;
	run;

	data _null_;
		set __RST end=stop;
		if _n_=1 then call execute("proc datasets lib=work nolist nodetails nowarn nofs;modify &outset;");
		%if %upcase(&type) eq XLS2SAS %then %do;
			call execute(" rename _c" || strip(_n_)     ||"= "|| strip(ddeName) ||";");
			call execute(" label    " || strip(ddeName) ||"= "|| strip(lstLabel)||";");
		%end;%else %do;
			call execute(" rename _c" || strip(_n_)     ||"= "|| strip(lstName) ||";");
			call execute(" label    " || strip(lstName) ||"= "|| strip(lstLabel)||";");
		%end;;
		if stop then call execute("quit;");
	run;
	proc sql; drop table __rst; quit;

	options source2 source;
%mend;

