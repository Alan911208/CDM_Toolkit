/**************************************************************
* Program:      dReview
* Function:     Data Review;
* Date:         Oct 10, 2012
* Programmer:   Fangqiu Xiao
*
* ClinStReport Applications for Clinical Data Management
* 
* Sample Usage:
* %dReview(path=P:\D589BL00022\RawData);
* %dReview(path=P:\D589BL00022\RawData,print=Simple);

* All Right Reserved 
*
* Change history:
*
***************************************************************/

*****以下代码请不要修改*************;
%macro dReview(path=D:\Temp\tmp,print=Original,showFail=NO);;

libname dReview "&Path.";
options nofmterr fmtsearch=(dReview.formats) noxsync noxwait xmin;;
filename sas2xl dde 'excel|system';
filename topic_ dde 'excel|system!topics' lrecl=5000;

data _null_;
	length fid rc start stop time 8;
	fid=fopen('sas2xl','s');
	if (fid le 0) then do;
		rc=system('start excel');
		start=datetime();
		stop=start+10;
		do while (fid le 0);
			fid=fopen('sas2xl','s');
			time=datetime();
			if (time ge stop) then fid=1;
		end;
	end;
	rc=fclose(fid);
run;

data _tpc0_;
	length topic_ $100;
	infile topic_ pad dsd notab dlm='09'x;
	input topic_ $ @@;
	topic_=scan(topic_,2,"]");
	if ^missing(topic_) and topic_ ne ":" then output;
run;

data _null_;
	file sas2xl;
	put '[workbook.next()]';
	put '[workbook.insert(3)]';
run;
data _tpc1_;
	length topic_ $100;
	infile topic_ pad dsd notab dlm='09'x;
	input topic_ $ @@;
	topic_=scan(topic_,2,"]");
	if ^missing(topic_) and topic_ ne ":" then output;
run;
proc sort data=_tpc1_; by topic_;run;

data _null_;
	set  _tpc1_ end=stop;
	if topic_="宏1" or upcase(topic_)="MACRO1" then call symput("mSheet",trim(topic_));
	if stop then call symput("mSheetn",compress(_n_-1));
run;
filename xlmacro dde "excel|&mSheet.!r1c1:r100c1" notab;

%local exData;
%let exData="WVISIT","LABREF","CMEDDRA","CUSERD","MISPAGE","CWHODD","PAGESTATUS","CROSSCHK","TRANSLATION","QUERY","FORMATS","PAT";
proc sql noprint;
	create table toc as select memname,memlabel as desc
	from sashelp.vtable where upcase(libname)=upcase("dReview") and memtype="DATA" and memname ^in ( &exData ) order by memname;
	select distinct memname into:datalist separated by " "
	from sashelp.vtable where upcase(libname)=upcase("dReview") and memtype="DATA" and memname ^in ( &exData ) order by memname ;
quit;;
data TOC(keep=memname desc);
	length  memname $40 desc $200;
	set TOC;
	if desc eq '' then desc=memname;
	label memname='Sheet Name' desc='Descriptions';
run;
%if %sysfunc(exist(dReview.query,data)) %then %do;
data crosschk0(drop=tblname qrytype username);
	length var $32 line 8;
	set %if %sysfunc(exist(dReview.crosschk,data)) %then dReview.crosschk; dReview.query(in=a);
	var=upcase(var);
	if a then do;
		dataset=tblname;
		line=compress(var,compress(var,'0123456789'));
		if missing(line) then line=1;
		var=compress(var,'0123456789');
		if qrytype="ENTRYqry" and username="Pending" then finding='025';
		if qrytype="EDITqry"  and username="Pending" then finding='024';
		if qrytype="ENTRYqry" and username="Closed"  then finding='023';
		if qrytype="EDITqry"  and username="Closed"  then finding='022';
		if qrytype ^in ("EDITqry","ENTRYqry")        then finding='021';
	end;
	select (finding);
		when ('Data Confirmed') finding='001';
		when ('Supported by AE/SAE record') finding='002';
		when ('Supported by lab finding') finding='003';
		when ('Supported by medical history record') finding='004';
		when ('Supported by operation history record') finding='005';
		when ('Supported by physical examination finding') finding='006';
		when ('Supported by treatment/drugs record') finding='007';
		when ('Supported by vital sign finding') finding='008';
		when ('021','022','023','024','025')  finding= finding;
		otherwise finding='999';
	end;
run;
%end;
%if ^%sysfunc(exist(dReview.query,data)) %then %do;
	proc sql noprint;
	create table Crosschk0 ( studyId Char (20) "Study Identifier",SiteId char (10) "Study Site Identifier ",usubjid char (10) "Unique Subject Identifier",
	             init char (4) "Subject Initials ",PageID char (6) "Page Identifier ",dataset char(20) "Dataset Name",var char(32) "Variable ",
	             line num "Line Identifier",_Find char (200) "Findings ",Finding char (200) "Findings ");
	             insert into  Crosschk0 (studyid) values('-----');
	quit;
%end;


proc sort data=crosschk0; by usubjid pageid dataset var line finding; run;

data crosschk0;
	length _find $200;
	retain _find '';
	set crosschk0; 
	by usubjid pageid dataset var line finding;
	if first.line then _find="";
		_find=left(right(_find)||" "||left(finding));
	if last.line then do;
		finding=_find;output;
	end;
run;
%let kk=1;
%do %while (%scan(TOC &datalist, &kk) ne );
	data _null_;
		file sas2xl;
		put '[workbook.next()]';
		put '[workbook.insert(1)]';
	run;
	data _tpc2_;
		length topic_ $100;
		infile topic_ pad dsd notab dlm='09'x;
		input topic_ $ @@;
		topic_=scan(topic_,2,"]");
		if ^missing(topic_) and topic_ ne ":" then output;
	run;
	proc sort data=_tpc2_; by topic_;run;
	data _null_;
		merge _tpc1_(in=a) _tpc2_(in=b);
		by  topic_ ;
		if ^a and b then call symput("ShtName",trim(topic_));
	run;
	proc sql;
		drop table _tpc2_;
	quit;

	data _null_;
		file xlmacro;
			put '=workbook.name("' "&ShtName" '","' "%scan(TOC &datalist, &kk)" '")';
			put '=halt(true)';
			put '!dde_flush';
			file sas2xl;
			put '[run("'"&mSheet."'!r1c1")]';
	run;



	%if %scan(TOC &datalist, &kk) ne TOC %then %do;
		data %scan(TOC &datalist, &kk);
			set dReview.%scan(TOC &datalist, &kk)
			%if %upcase(&print)=SIMPLE %then (drop=studyid domain subjid siteid random visitnum);;
		run;
		proc sort data= %scan(TOC &datalist, &kk);
			by usubjid pageid line;
		run;
	%end;

	%sas2xls(inset=%scan(TOC &datalist, &kk),outsheet=%scan(TOC &datalist, &kk));

	%if %sysfunc(exist(dReview.query,data)) %then %do;
	%if %scan(TOC &datalist, &kk) eq TOC %then %do;
		filename xlTOC dde "excel|TOC!r1c4:r20c5" notab;
		data _null_;
			file xlTOC;
			put "Lengend";
			put '09'x 'Data Confirmed';
			put '09'x 'Supported by AE/SAE record';
			put '09'x 'Supported by lab finding';
			put '09'x 'Supported by medical history record';
			put '09'x 'Supported by operation history record';
			put '09'x 'Supported by physical examination finding';
			put '09'x 'Supported by treatment/drugs record';
			put '09'x 'Supported by vital sign finding';
			put '09'x 'Unknown Type';put;
			put '09'x 'Entry Comment';put;
			put '09'x 'Answered programme-check-query';put;
			put '09'x 'Answered key-in-Query';put;
			put '09'x 'Not Answered programme-check-query';put;
			put '09'x 'Not Answered key-in-query';
		run;

		data _null_;
			file sas2xl;
			put '[select("R2C4:R2C4")]'; put '[patterns(1,,07)]';
			put '[select("R3C4:R3C4")]'; put '[patterns(1,,19)]';
			put '[select("R4C4:R4C4")]'; put '[patterns(1,,24)]';
			put '[select("R5C4:R5C4")]'; put '[patterns(1,,34)]';
			put '[select("R6C4:R6C4")]'; put '[patterns(1,,36)]';
			put '[select("R7C4:R7C4")]'; put '[patterns(1,,38)]';
			put '[select("R8C4:R8C4")]'; put '[patterns(1,,39)]';
			put '[select("R9C4:R9C4")]'; put '[patterns(1,,40)]';
			put '[select("R10C4:R10C4")]'; put '[patterns(1,,03)]';
			put '[select("R12C4:R12C4")]'; put '[border(5,,,,,,28)]';
			put '[select("R14C4:R14C4")]'; put '[border(2,,,,,,01)]';
			put '[select("R16C4:R16C4")]'; put '[border(6,,,,,,01)]';
			put '[select("R18C4:R18C4")]'; put '[border(2,,,,,,03)]'; 
			put '[select("R20C4:R20C4")]'; put '[border(6,,,,,,03)]';
		run;
		filename xlTOC clear;
	%end;
	%end;

	proc sql noprint;
		select count(*) into:go_on from crosschk0 where dataset="%scan(TOC &datalist, &kk)";
	quit;
	%if &go_on>0 %then %do;
		proc sort data=crosschk0(where=(dataset="%scan(TOC &datalist, &kk)")) out=crosschk;
			by var usubjid pageid line;
		run;
		proc sql noprint;
			create table _VARlst as select upcase(name) as var, varnum as _col from sashelp.vcolumn 
			where upcase(libname)="WORK" and memtype="DATA" and memname="%scan(TOC &datalist, &kk)" order by upcase(name);
		quit;
		%if %upcase(&showFail) eq YES %then %do;
		data __%scan(TOC &datalist, &kk)(drop=_col dataset _find);;
			length STUDYID $20 DOMAIN $8 USUBJID $20 SITEID $10 INIT $4;
			length PAGEID $6 LINE 8 VAR $32 FINDING $200;
			merge crosschk(in=a)  _VARlst(in=b);
			by var;
			DOMAIN=dataset;
			if a and ^b;
			select (finding);
				 when('001') finding='Data Confirmed';
				 when('002') finding='Supported by AE/SAE record';
				 when('003') finding='Supported by lab finding';
				 when('004') finding='Supported by medical history record';
				 when('005') finding='Supported by operation history record';
				 when('006') finding='Supported by physical examination finding';
				 when('007') finding='Supported by treatment/drugs record';
				 when('008') finding='Supported by vital sign finding';
				 when('999') finding='Unknown Type';
				 when('021') finding='Entry Comment';
				 when('022') finding='Answered programme-check-query';
				 when('023') finding='Answered key-in-Query';
				 when('024') finding='Not Answered programme-check-query';
				 when('025') finding='Not Answered key-in-query';
				 otherwise finding='Error';
			end;
		run;
		proc sql noprint;
			select nobs into : myOBS0 from sashelp.vtable 
			where upcase(libname)="WORK" and memtype="DATA" and memname="__%scan(TOC &datalist, &kk)";
		quit;
		%if &myOBS0>0 %then %put WA%str()RNING: There are %cmpres(&myOBS0) items not marked in %scan(TOC &datalist, &kk);
		%else %do;
			proc sql noprint;
				drop table __%scan(TOC &datalist, &kk);
			quit;
		%end;
		%end;
		data crosschk(drop=_col);
			length col $8;
			merge crosschk(in=a)  _VARlst(in=b);
			by var;
			col=compress("__c"||_col);
			if a and b;
		run;
		proc sort data=crosschk;by usubjid pageid line;run;
		proc transpose data=crosschk out=crosschk(drop=_name_ _label_);
			id col;
			var finding;
			by usubjid pageid line;
		run;

		data crosschk;
			merge %scan(TOC &datalist, &kk) (IN=AAA) crosschk;
			by  usubjid pageid line;
			IF AAA;
		run;
		data _null_;
			length range $50;
			file sas2xl;
			set crosschk;
			by  usubjid pageid line;
			array _chr_ _CHAR_;
			do over _chr_;
				if substr(vname(_chr_),1,3)="__c" and ^missing(_chr_) then do;
					range=compress("r"||_n_+1||substr(vname(_chr_),3)||":r"||_n_+1||substr(vname(_chr_),3));
					put '[select("' range '")]'; 
					i=1;
					do while (scan(_chr_, i) ne "");
						select (scan(_chr_, i));							
						    when ('001') put '[patterns(1,,07)]';
						    when ('002') put '[patterns(1,,19)]';
						    when ('003') put '[patterns(1,,24)]';
						    when ('004') put '[patterns(1,,34)]';
						    when ('005') put '[patterns(1,,36)]';
						    when ('006') put '[patterns(1,,38)]';
						    when ('007') put '[patterns(1,,39)]';
						    when ('008') put '[patterns(1,,40)]';
						    when ('999') put '[patterns(1,,03)]';
							when ('021') put '[border(5,,,,,,28)]';
							when ('022') put '[border(2,,,,,,01)]';
							when ('023') put '[border(6,,,,,,01)]';
							when ('024') put '[border(2,,,,,,03)]';
							when ('025') put '[border(6,,,,,,03)]';
						end;
						i=i+1;
					end;
				end;
			end;
		run;
	%end;

	data _tpc1_;
		length topic_ $100;
		infile topic_ pad dsd notab dlm='09'x;
		input topic_ $ @@;
		topic_=scan(topic_,2,"]");
		if ^missing(topic_) and topic_ ne ":" then output;
	run;
	proc sort data=_tpc1_; by topic_;run;
	proc sql;
		drop table %scan(TOC &datalist, &kk);
	quit;
	%let kk=%eval(&kk+1);
%end;

data _null_;
	set _tpc0_ end=stop;
	call symput("sht"||compress(_n_),trim(topic_));
	if stop then call symput("mSheetn",compress(_n_));
run;
%do kk=1 %to &mSheetn;
	data _null_;
		file xlmacro;
			put '[error(false)]';
			%if &kk=1 %then put '=workbook.hide("' "&mSheet" '")';;
			put '=workbook.delete("' "&&&sht&kk" '")';;
			put '=halt(true)';;
			put '!dde_flush';
			file sas2xl;
			put '[run("'"&mSheet."'!r1c1")]';
	run;
%end;
filename sas2xl clear;
filename topic_ clear;
filename xlmacro clear;

proc sql;
	drop table _tpc0_,_tpc1_ ,crosschk0 %if %sysfunc(exist(_varlst,data)) %then ,_varlst;
			   %if %sysfunc(exist(crosschk,data)) %then ,crosschk;;
quit;

%mend;
