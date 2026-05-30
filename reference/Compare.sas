
/**************************************************************
* Program:      Compare
* Function:     Compare previus data to current data and report to Excel;
* Date:         May 07, 2017
* Programmer:   Fangqiu Xiao
*
* Clinflash Applications for Clinical Data Management
* 
* Sample Usage:

  %Compare(
	pCompare=D:\temp\tmp\old,
	pBase=D:\temp\tmp\new,
	kVar=studyId __STUDYEVENTREPEATKEY  usubjid visit pageid PAGE form line);

	mode=1:忽略前导空格
	mode=2:忽略大小写
	mode=3:忽略前导空格以及大小写


* All Right Reserved 
*
* Change history:
* Nov 9： Display  non-empty original data when record is only in NEW or OLD; 
* 17Jan2018: Fix the data export to cover other sheet bugs.
* 1July2021; Auto data format translate

*俞涛20240913：新增eVar变量用于排除不需要比对的变量
*俞涛20260519：create table toc as select memname,libname,memlabel as desc改为create table toc as select memname,libname,memlabel as desc length=200，提前定义长度，否则在日志中会报数据截断的提醒

***************************************************************/

%macro Compare(pCompare=,pBase=,kVar=,eVar=,mode=9);
	libname new "&pBase";
	libname old "&pCompare";

	options fmterr noxsync noxwait xmin nosource nonotes nomprint missing='';;
	options fmterr noxsync noxwait xmin source notes mprint;;
	%let _tblALL=%str();
	proc sql noprint;
		create table toc as select memname,libname,memlabel as desc length=200
		from sashelp.vtable where upcase(libname) in("NEW","OLD") and memtype="DATA"  order by memname;
		select distinct memname into:_tblALL separated by " "
		from sashelp.vtable where upcase(libname) in("NEW","OLD") and memtype="DATA"  order by memname;
	quit;

	proc sort data=toc; by memname libname; run;

	data TOC(keep=memname location desc _recNote);
		length  memname $40  desc $200 location $100 _recNote $50;
		set TOC;by memname libname;
		location="Double data sets";
		if first.memname and last.memname then do;
			if libname="NEW" then do; location="Single[&pBase]"; _recNote="New Only";end;
			if libname="OLD" then do; location="Single[&pCompare]"; _recNote="Old Only";end;
		end;
		if last.memname;
		label memname='SheetName' location='Location' desc='Descriptions' _recNote="No. of different datapoints";
	run;

	%let k=1;
	data _varlst;
		length var $32;
		%do %while (%scan(&kVar,&k,%str( )) ne );
			var=upcase("%scan(&kVar,&k,%str( ))"); kvar=1;output;
			%let k=%eval(&k+1);
		%end;
	run;
	proc sort data=_varlst; by var; run;

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
  

	%let kk=1;
	%do %while (%scan(&_tblALL, &kk) ne );
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
			merge _tpc1_(in=a) _tpc2_(in=b where=(topic_ ne 'TOC'));
			by  topic_ ;
			if ^a and b then call symput("ShtName",trim(topic_));
		run;

		proc sql;
			drop table _tpc2_;
		quit;

		data _null_;
			file xlmacro;
				put '=workbook.name("' "&ShtName" '","' "%scan(&_tblALL, &kk)" '")';
				put '=halt(true)';
				put '!dde_flush';
				file sas2xl;
				put '[run("'"&mSheet."'!r1c1")]';
		run;

		%let mySet=%upcase(%scan(&_tblALL, &kk));
		proc sql noprint;
			create table _New as select upcase(name) as Var, varnum as nNew, type as nType, length as nlen,label as nlbl,format as nfmt
			from sashelp.vcolumn  where upcase(libname)="NEW"  and memtype="DATA" and memname="&mySet" order by Var;
			create table _Old as select upcase(name) as Var, varnum as nOld ,type as oType, length as olen,label as olbl,format as ofmt
			from sashelp.vcolumn  where upcase(libname)="OLD"  and memtype="DATA" and memname="&mySet" order by Var;
		quit;

		data _varlst1;
			length co $20;
			merge  _Old(in=ooo)  _New(in=nnn)  _varlst;	by Var;
			if oType ne nType then kvar=.;
			if ooo and ^nnn then co='Not Reported';;
			if ^ooo and nnn then co='New Variable';;
/*2024-09-24 update for the english sutdy*/
			nlbl=translate(nlbl,byte(39),byte(34));;
            olbl=translate(olbl,byte(39),byte(34));;

			if Var='' then delete;
		run;
	    proc sql noprint;
			select count(*) into:kkk from _varlst1 where ^missing(nold*nnew*kvar);
		quit;

		data _bbbb_(drop=i name_a);
		length memnamea Var $200 ;
		name_a=upcase("&EVar");
		memnamea=upcase("&mySet");
  		do i=1 by 1 until(Var=' ');
     		Var=scan(name_a,i,' ') ;
     		if i=1 or Var ne ' ' then output;
  		end;
		run;

proc sort data=_varlst1;  by Var ; run;
proc sort data=_bbbb_;  by Var ; run;

		data _varlst1;
		merge  _bbbb_(in=a) _varlst1(in=b);
		by Var;
		if a and b then kvar=2;

		if Var='' then delete;
		run;

		proc sort data=_varlst1;  by descending nNew nold ; run;
		data _varlst1;
			set _varlst1;
			by descending nNew nold ;
			if nnew ne . then vnum=nnew;else  vnum=_n_;
			if kvar=1 or kvar=2 then cvar=var;else cvar="__V"||compress(vnum);
			if (kvar=2 and otype eq 'num') then cvar="_"||compress(var);
		run;
		data _pppp_(keep=memnamea cvar);set _varlst1;if kvar=2 then output;run;
%let mmmm=0;
data _null_;
set _pppp_ nobs=nobs;
if nobs gt 0 then call symput("mmmm","1");
run;
%put mmmm=&mmmm;
		proc sort data=_varlst1 ;by vnum; run;

		%if %eval(&k-1) eq &kkk %then %do;
			data _null_;
				set  _varlst1 end=stop;by vnum;
				if missing(nlbl) then nlbl=olbl;
				if kvar eq 1 then do;
					if _n_ eq 1 then call execute("proc sql noprint;create table _itmp_ ("||var||" " ||ntype ||"("||nlen||') label="'||strip(nlbl)||'");');
					if _n_ ne 1 then call execute("alter table _itmp_ add "||var||" " ||ntype ||"("||nlen||') label="'||strip(nlbl)||'";');
					if  ^missing(nfmt) then call execute("alter table _itmp_ modify "||var||" " ||ntype ||"("||nlen||') format='||nfmt||';');
				end; else do;
					if _n_ eq 1 then call execute("proc sql noprint;create table _itmp_ ("||cvar||' char(2000) label="'||strip(nlbl)||'");');
					if _n_ ne 1 then call execute("alter table _itmp_ add "||cvar||' char(2000) label="'||strip(nlbl)||'";');
					call execute("alter table _itmp_ add V"||cvar||' char(2000) label="'||strip(nlbl)||'";');
					call execute("alter table _itmp_ add VV"||cvar||' char(2000) label="'||strip(nlbl)||'";');
				end;
				if stop then call execute("quit;");
			run;
 
			proc sort data=new.&mySet out=_new&mySet; by &kVar; run;
			proc sort data=old.&mySet out=_old; by &kVar; run;


			data _null_;
				set  _varlst1(where=(kvar ne 1 and otype eq 'char')) end=stop ;
				if _n_=1 then call execute("proc datasets lib=work nolist;modify _old; rename " || var ||"= V"|| compress(cvar)||";");
				if _n_ ne 1 then call execute("modify _old; rename " || var ||"= V"|| compress(cvar)||";");
				if stop then call execute("quit;");
			run;
			data _null_;
				set  _varlst1(where=(kvar ne 1 and otype eq 'num')) end=stop ;
				if _n_=1 then call execute("data _old; set _old;");
				if ^missing(ofmt) then call execute("V"|| compress(cvar)||"=put("||var||","||ofmt||");");
				if  missing(ofmt) then call execute("V"|| compress(cvar)||"=strip("||var||"||' ');");
				call execute("drop "|| var ||";");
				if stop then call execute("run;");
			run;
			data _null_;
				set  _varlst1(where=(kvar ne 1 and kvar ne 2 and ntype eq 'char')) end=stop ;
				if _n_=1 then call execute("proc datasets lib=work nolist;modify _new&mySet; rename " || var ||"= "|| compress(cvar)||";");
				if _n_ ne 1 then call execute("modify _new&mySet; rename " || var ||"= "|| compress(cvar)||";");
				if stop then call execute("quit;");
			run;
			data _null_;
				set  _varlst1(where=(kvar ne 1 and ntype eq 'num')) end=stop ;
				if _n_=1 then call execute("data _new&mySet; set _new&mySet;");
				if ^missing(nfmt) then call execute(compress(cvar)||"=put("||var||","||nfmt||");");
				if  missing(nfmt) then call execute(compress(cvar)||"=strip("||var||"||' ');");
				call execute("drop "|| var ||";");
				if stop then call execute("run;");
			run;

            %let nnnn=0;
			data _itmp_(drop=iiii zzzz);
			   retain _nnnn 0;
				merge _itmp_ _new&mySet(in=_dNew) _old(in=_dOld) end=stop; by &kVar; 
				if ^_dOld then _recNEW=-1;else _recNew=1;
				array s _char_;
				do iiii=1 to dim(s)-2;
					if length(vname(s(iiii)))>3 then do;
					if substr(vname(s(iiii)),1,3)="__V" then do;
						zzzz=' ';
						%if &mode=1 %then if strip(s(iiii)) eq strip(s(iiii+1)) then do;
						%else %if &mode=2 %then if upcase(s(iiii)) eq upcase(s(iiii+1)) then do;
						%else %if &mode=3 %then if upcase(strip(s(iiii))) eq upcase(strip(s(iiii+1))) then do;
						%else if  s(iiii)=s(iiii+1) then do;;
							if missing(vformat(s(iiii))) then s(iiii+2)=s(iiii);else s(iiii+2)=putc(s(iiii),vformat(s(iiii)));
						end;
							else do;
							if missing(vformat(s(iiii))) then s(iiii+2)=s(iiii);else s(iiii+2)=putc(s(iiii),vformat(s(iiii)));
							if missing(vformat(s(iiii+1))) then do;
								if _recNew=-1 then s(iiii+2)=strip(s(iiii+2));
								else if ^_dNew  then s(iiii+2)=strip(s(iiii+1));
								else s(iiii+2)=strip(s(iiii+2))||"/"||strip(s(iiii+1));
							end;
							else do;
                               if _recNew=-1 then s(iiii+2)=strip(s(iiii+2));
							   else if ^_dNew  then s(iiii+2)=strip(putc(s(iiii+1),vformat(s(iiii+1))));
							   else s(iiii+2)=strip(s(iiii+2))||"/"||strip(putc(s(iiii+1),vformat(s(iiii+1))));

							end;
							zzzz='1';
							_nnnn =_nnnn +1;
						end;
						if zzzz='1' and _recNew>0 then _recNew=_recNew+1;
						s(iiii+1)=zzzz;
					end;
					end;
				end;
			length _recNote $50.;
				if _recNew=-1 then _recNote="New     ";
				if _recNew> 1 then _recNote="Changed ";
				if ^_dNew  then _recNote="Old  "; 
				if stop then call symput ("nnnn",_nnnn);
				label _recNote='Notes';
			run;



%if &mmmm=1 %then %do;
proc sort data=_pppp_;by memnamea;run;
data _ffff_(keep=Evar Evar_A);
set _pppp_;
by memnamea;
length  Evar Evar_A $2000.;
retain Evar Evar_A;
if first.memnamea then do;Evar=" ";Evar_A=" ";end;
Evar=strip(strip(Evar)||" VV"||strip(cvar)||"="||strip(cvar)||";");
Evar_A=strip(strip(Evar_A)||" if "||strip(cvar)||" ^=V"||strip(cvar)||" then VV"||strip(cvar)||"=catx(""/"","||strip(cvar)||",V"||strip(cvar)||");");
if last.memnamea;
run;

data _null_;
set _ffff_;
f='data _itmp_;set _itmp_;'||strip(Evar)||strip(Evar_A)||'run;';
call execute(f);
run;
%end;

			%let vkeep=;
			%let vkeep0=;
			proc sql noprint;
				select "vv"||cvar into: vkeep separated by ' '
				from _varlst1 where kvar ne 1;
				select "v"||cvar into: vkeep0 separated by ' '
				from _varlst1 where kvar ne 1;
			quit;
			data _new&mySet;
				set _itmp_(keep=&vkeep &kvar _recNote);
			run;

			data _old;
				set _itmp_(keep=&vkeep0 &kvar _recNote);
			run;

			%if &nnnn ne 0 %then %do;
	            data toc;
					set toc;
					if memname="%upcase(%scan(&_tblALL, &kk))" then _recNote="&nnnn";
					label _recNote="No. of different datapoints";
				run;	
			%end;

			%sas2xls(inset=_new&mySet,outsheet=&mySet);

			filename sas2xl dde 'excel|system';

			data _Null_;
				file sas2xl ;
				length jjj $200;
				set _old;
				array s _char_;
				do kkkk=1 to dim(s);
					if length(vname(s(kkkk)))>4 then do;
					if substr(vname(s(kkkk)),1,4)="V__V" and s(kkkk)="1" then do;
						jjj="R"||compress(_N_+1)||"C"||compress(substr(vname(s(kkkk)),5))||":R"||compress(_N_+1)||"C"||compress(substr(vname(s(kkkk)),5));
						put '[select("' jjj '")]'; 
						put '[patterns(1,,07)]';
					end;
					end;
				end;
				array ss _numeric_;
		        if  _recNote="New     " then do;
					jjj="R"||compress(_N_+1)||"C1:R"||compress(_N_+1)||'C'||compress(dim(s)+dim(ss)-2); 
					put '[select("' jjj '")]'; 
					put '[patterns(1,,4)]';
				end;
		        if  _recNote="Old     " then do;
					jjj="R"||compress(_N_+1)||"C1:R"||compress(_N_+1)||'C'||compress(dim(s)+dim(ss)-2); 
					put '[select("' jjj '")]'; 
					put '[patterns(1,,15)]';
				end;
			run;

		%end;%else %do;
			data _new&mySet;
				set %if %sysfunc(exist(new.&mySet,data)) %then new.&mySet;%else old.&mySet; end=stop;length _recNote $50.;
				%if %sysfunc(exist(new.&mySet,data)) %then _recNote="New"; %else _recNote="Old";;
				if stop then do;
					array s _CHAR_;
					array nn _numeric_;
					if stop then call symput("nvars",dim(s)+dim(nn));
				end;
				label _recNote='Notes';
			run;
			title %sysfunc(pathname(mymacros));;
			%sas2xls(inset=_new&mySet,outsheet=&mySet);

			filename sas2xl dde 'excel|system';

			data _Null_;
				file sas2xl ;
				length jjj $200;
				set _new&mySet;
				jjj="R"||compress(_N_+1)||"C1:R"||compress(_N_+1)||'C'||compress("&nvars"); 
				put '[select("' jjj '")]'; 
				%if %sysfunc(exist(new.&mySet,data)) %then put '[patterns(1,,4)]';;
				%if ^%sysfunc(exist(new.&mySet,data)) %then put '[patterns(1,,15)]';;
			run;
		    proc sql;
				drop table _new

%if %sysfunc(exist(_old,data)) %then ,_old;;
			quit;
		%end;

	data _tpc1_;
		length topic_ $100;
		infile topic_ pad dsd notab dlm='09'x;
		input topic_ $ @@;
		topic_=scan(topic_,2,"]");
		if ^missing(topic_) and topic_ ne ":" then output;
	run;
	proc sort data=_tpc1_; by topic_;run;

		%let kk=%eval(&kk+1);
	%end;
	data _null_;
		set _tpc0_ end=stop;
		if _n_=1 then call symput("ShtName",trim(topic_));
		call symput("sht"||compress(_n_-1),trim(topic_));
		if stop then call symput("mSheetn",compress(_n_-1));
	run;

	data _null_;
		file xlmacro;
		put '=workbook.name("' "&ShtName" '","' "TOC" '")';
		put '=halt(true)';
		put '!dde_flush';
		file sas2xl;
		put '[run("'"&mSheet."'!r1c1")]';
	run;

	filename xlTOC dde "excel|TOC!r1c1:r1c1" notab;
	data _null_;
		file xlTOC;
		put "List of Data sets";
	run;

	%sas2xls(inset=TOC,outsheet=TOC,stRow=2);

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
		drop table _tpc0_,_tpc1_,_varlst,toc;
	quit;
%mend;
