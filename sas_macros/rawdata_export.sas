
/**********************************************************************************
*  Company XXX
*  Program name     : rawdata_content_report.sas
*  Project          : for all project
*  Written by       : Weineng zhou
*  Date of creation : 2022-04-29
*  Description      : review raw data
*  Macros called    : %currentroot, %get_form, %Report
*  Input file       : all rawdata
*  Output file      : &projroot.doc\programming\rawdata_content_report.xlsx
*  Revision History : 
*  Date      Author       Description of the change
**********************************************************************************/


dm "log;clear;";
proc datasets lib=work kill nolist memtype=data;run;
quit;
%macro currentroot;
%global currentroot pgmname currentpath;
%let currentroot= %sysfunc(getoption(sysin));
%if "&currentroot" eq "" %then %do;
%let currentroot= %sysget(SAS_EXECFILEPATH);
%end;
%let pgmname=%scan(&currentroot, -1, \);
%let currentpath = %sysfunc(prxchange(s/(.*)\\.*/\1/, -1, &currentroot));
%mend;
%currentroot;
%put &currentpath.;

option validvarname=upcase;
libname raw "&currentpath.";


proc contents data=raw._all_ out=content noprint;
run;
proc sql noprint;
	select distinct memname into :ds_list separated by ' '
	from content;
	select distinct memname into :rawdata_list separated by '", "'
	from content;
	;
quit;
%put &ds_list.;
%put "&rawdata_list.";

*copy dataset from raw to work library;
proc datasets library=raw nolist;
    copy in=raw out=work;
    select &ds_list.;
quit;
%put &ds_list.;


%macro allvar_length_optimize(lib=work);

%local _obs datasetlabel n_char allvar i nd;

%* %do %while ;
%let x = 1;
%let ds = %scan(&ds_list, &x, %str( )); %put &ds.;

%do %while (&ds ne );

	proc sql noprint;
		select count(*) into :_obs
		from &lib..&ds.;
	quit;
	%put &_obs;

	%if %eval(&_obs>0) %then %do;

	proc contents data=work.&ds. out=a_content varnum noprint;
	run;

	data _null_;
		set a_content end=eof;
		if eof then call symputx("datasetlabel",memlabel);
	run;		
	%put &datasetlabel;

	proc sql noprint;
	    select compress(put(count(*),best.)) into :n_char from a_content where type=2;
	    select name into :oriname1-:oriname&n_char from a_content where type=2;
		select name into :allvar separated by ' ' from a_content order by varnum;
	quit;

	proc sql;
		create table max_length as
			select 
				%do i=1 %to &n_char;
					max(length(&&oriname&i)) as &&oriname&i
					%if &i=&n_char %then %str();%else %str(,);
				%end;
			from &lib..&ds.;
		;
	quit;

		proc transpose data=max_length out=max_length_verti;
			var _all_;
		run;

		data _null_;
		    set max_length_verti;
		    call symputx(compress("charlen"||put(_N_,best.)),_name_||" $"||strip(put(col1,best.)));
		run;
		options varlenchk=nowarn;

		data &lib..&ds. %if %length(&datasetlabel) %then (label="&datasetlabel");;
			retain &allvar;
			length 
				%do i=1 %to &n_char;
					&&charlen&i
				%end;
			;
			set &lib..&ds.;
			informat _character_;
			format _character_;
		run;
	%*%symdel _obs datasetlabel n_char allvar i nd; 
	%*%if &_obs %then %do;
	%end; 
%let x = %eval(&x + 1);
%let ds = %scan(&ds_list, &x, %str( ));
%end;

%mend;
%allvar_length_optimize;


proc contents data=work._all_ out=content noprint;
run;
proc sort data=content;
	by memname varnum;
run;

data content;
	set content;
	if memname in ("&rawdata_list.");
	if not prxmatch("/__|_U/",name);
	variable_name=name;
	if prxmatch("/Specify/i",label) then varibale_label=strip(scan(label,1,'('))||"(Specify)";
	else if prxmatch("/AETERM|CMTRT|PRTRT|PR\dTRT|ATC/i",name) then varibale_label=strip(scan(label,1,'('))||"(Specify)";
	else varibale_label=strip(scan(label,1,'('));
	if type=2 then vartype="Char";
	else vartype="Num";
	keep memname variable_name varibale_label vartype length varnum;
	proc sort ;
	by memname varnum;
run;


data rawdata_list;
	set content;
	proc sort nodupkey;
	by memname;
run;
data _null_;
	set rawdata_list end=total;
	call symputx("rawdata"||strip(put(_N_,best.)), memname);
	if total then call symputx('maxiter', _N_);
run;


%macro get_form;

%let k = 0;
%do %while(&k. < &maxiter. ); %put &k.;
%let k = %eval(&k+1);

	data _null_;
		set work.&&rawdata&k.;
		call symputx("FORM", FORM);
	run;

	data _null_;
		if 0 then set work.&&rawdata&k. nobs=nobs;
		if nobs=0 then call symputx("FORM", "");
	run;

	%put &FORM.;

	data toc__&&rawdata&k.;
		rawdata = "&&rawdata&k."; 
		FORM = "&FORM.";
	run;

%end;

%mend;
%get_form;


data toc;
	length rawdata form $200;
	set toc__:;
run;


options papersize=letter orientation=landscape nodate nonumber center missing=" " nobyline; 
options formchar="|----|+|---+=|-/\<>*"; 
ods escapechar="@";

ods html close;
ods excel file="&currentpath.\rawdata_content.xlsx" style=normal;
%macro Report;

ods excel options(sheet_name="CONTENT" absolute_row_height='16' absolute_column_width='10,40,10' autofilter = 'column');

proc report data=toc missing center nowd headline headskip split = '~'

	style(header)={just=left asis=on nobreakspace=on color=black backgroundcolor=#DCE6F1 font=("Arial", 10pt) font_weight=bold font_size=2 };

	column rawdata form ;

	define rawdata /display	'rawdata';
	define form    /display	'form';

run;


%let k = 0;
%do %while(&k. < &maxiter. ); %put &k.;
%let k = %eval(&k+1);

ods excel options(sheet_name="%upcase(&&rawdata&k.)" absolute_row_height='16' absolute_column_width='15,50,8,8,20,10,5,20,10,12' autofilter = 'column');

proc report data=content(where=(memname="%upcase(&&rawdata&k.)")) missing center nowd headline headskip split = '~' 

	style(header)={just=left asis=on nobreakspace=on color=black backgroundcolor=#DCE6F1 font=("Time New Roman/ËÎÌו", 10pt) font_weight=bold font_size=2 };

	column variable_name varibale_label vartype length varnum ;

	define variable_name/display   	            'Variable_Name';	
	define varibale_label/display	            'Variable_Label';
	define vartype/display   	                'Type';
	define length/display	                    'Length';
	define varnum/display	                    'Variable_Order';

run;

%end;

%mend;
%Report;
ods excel close;

