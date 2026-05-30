/**************************************************************
* Program:      CutData
* Function:     按指定日期提取数据;
* Date:         Apr 13,2017
* Programmer:   Fangqiu Xiao
*
* Clinflash Applications for Clinical Data Management
* 
* Sample Usage: 201/301
	 %CutData(
		pSource=D:\temp\tmp,
		pTarget=D:\temp\tmp\partdata,
		DateCut=20170228,
		vDateCut=SV1.SVSTDTC_YMD SV2.SVSTDTC_YMD SV3.SVSTDTC_YMD,
		others=ae.aestdtc_ymd ce.CESTDTC_YMD CM3.cmstdtc_ymd DS2.dsstdtc1_ymd ds3.DSSTDTC_YMD,
		datacopy=formats ds4 co,
		byVar=Usubjid Visit);
	 %CutData(
		pSource=D:\temp\tmp,
		pTarget=D:\temp\tmp\partdata,
		DateCut=20170228,
		vDateCut=SV1.SVSTDTC_YMD,
		others=ae.aestdtc_ymd ce.CESTDTC_YMD CM3.cmstdtc_ymd CM4.cmstdtc_ymd DS2.dsstdtc1_ymd SV3.svstdtc_ymd,
		datacopy=FORMATS CO,
		byVar=Usubjid Visit);
*
* All Right Reserved 
*
* Change history:
* 
*
***************************************************************/

%macro CutData(
	pSource=D:\temp\tmp,
	pTarget=D:\temp\tmp\partdata,
	DateCut=20170228,
	vDateCut=SV.SVSTDTC_YMD,
	others=ae.aestdtc_ymd  CM.cmstdtc_ymd,
	datacopy=FORMATS CO,
	byVar=Usubjid Visit);

	libname iData "&pSource";
	libname oData "&pTarget";

	%let _ex_=%str();
	%let k=1;
	%do %while (%scan(&others,&k,%str( )) ne );
		%let _ex_=&_ex_ %upcase("%scan(%scan(&others,&k,%str( )),1,.)"),;
		%let k=%eval(&k+1);
	%end;

	%let k=1;
	%do %while (%scan(&datacopy,&k,%str( )) ne );
		%let _ex_=&_ex_ %upcase("%scan(&datacopy,&k,%str( ))"),;
		%let k=%eval(&k+1);
	%end;

	%if %sysfunc(exist(oData._rpt_,data)) %then %do;
		proc sql;
			drop table oData._rpt_;
		quit;
	%end;
	%if %sysfunc(exist(_rpt_,data)) %then %do;
		proc sql;
			drop table _rpt_;
		quit;
	%end;
	%if %sysfunc(exist(_vst_,data)) %then %do;
		proc sql;
			drop table _vst_;
		quit;
	%end;

	proc sql noprint;
		select memname into :inList separated by ' ' 
		from sashelp.vtable where libname="%upcase(iData)" and memtype="DATA" 
        %if %length(&_ex_)>1 %then and memname ^in (&_ex_ "????");;
	quit;

	%let k=1;
	%do %while (%scan(&vDateCut,&k,%str( )) ne );
		%if ^%sysfunc(exist(_VST_,data)) %then %do;
			data _VST_(keep=&byvar _vst_);
				set iData.%scan(%scan(&vDateCut,&k,%str( )),1,.);
				_vst_=%scan(%scan(&vDateCut,&k,%str( )),2,.);
			run;
		%end; %else %do;
			data _VST_(keep=&byvar _vst_);
				set _vst_ iData.%scan(%scan(&vDateCut,&k,%str( )),1,.)(in=b);
				if b then _vst_=%scan(%scan(&vDateCut,&k,%str( )),2,.);
			run;
		%end;
		%let k=%eval(&k+1);
	%end;

	proc sort data=_vst_(where=(^missing(_vst_))) out=_SVLST_(keep=%scan(&byvar,2,%str( ))) nodupkey ; 
		by %scan(&byvar,2,%str( ));
	run;

	data _vst_;
		set _vst_;
		if _vst_< input("&DateCut",yymmdd8.);
	run;

	proc sort data=_vst_;by &byvar; run;

	%let k=1;
	%let bylist=;
	%do %while (%scan(&byvar,&k,%str( )) ne );
		%let bylist=&byList %upcase("%scan(&byvar,&k,%str( ))"),;
		%let k=%eval(&k+1);
	%end;

	%let k=1;
	%do %while (%scan(&inList,&k,%str( )) ne );
		%let  vars=0;
		proc sql noprint;
			select count(name) into: vars from dictionary.columns where 
			libname=%upcase("iData") and name in ( &bylist "????" ) and
			memname=%upcase("%scan(&inList,&k,%str( ))");
		quit;

		%if &vars>0 %then %do;

			proc sort data= iData.%scan(&inList,&k,%str( )) out=_tmp_; by &byvar; run;
			proc sort data=_tmp_  out=_tmp0_ nodupkey; by %scan(&byvar,2,%str( )); run;

			%let _msg_=;
			data _null_;
				merge _tmp0_(in=a) _SVLST_(in=b);
				by %scan(&byvar,2,%str( )); 
				if a and ^b then do;
					call symput("_msg_","(不全匹配)"||left(%scan(&byvar,2,%str( ))));
				end;
			run;

			data oData.%scan(&inList,&k,%str( ));
				merge _tmp_(in=a) _VST_(in=b);
				by &byVar;
				if a and b;
				drop _vst_;
			run;

			%if ^%sysfunc(exist(_rpt_,data)) %then %do;
				data _rpt_;
					length memname note $32;
					memname=%upcase("%scan(&inList,&k,%str( ))");
					note="按 Visit &_msg_";
					seq=1;
				run;
			%end;%else %do;
				data _rpt_;
					set _rpt_ end=stop;
					output;
					if stop then do;memname=%upcase("%scan(&inList,&k,%str( ))");
					note="按 Visit &_msg_";
					seq=1;output;end;
				run;
			%end;
		%end;
		proc sql; drop table _tmp_,_tmp0_;quit;
		%let k=%eval(&k+1);
	%end;
	proc sql; drop table _vst_,_svlst_;quit;


	%let k=1;
	%do %while (%scan(&others,&k,%str( )) ne );
		data oData.%scan(%scan(&others,&k,%str( )),1,.);
			set iData.%scan(%scan(&others,&k,%str( )),1,.);
			if %scan(%scan(&others,&k,%str( )),2,.)< input("&DateCut",yymmdd8.);
		run;

		%if ^%sysfunc(exist(_rpt_,data)) %then %do;
			data _rpt_;
				length memname note $32;
				memname=%upcase("%scan(%scan(&others,&k,%str( )),1,.)");
				note="按 %scan(%scan(&others,&k,%str( )),2,.) ";
				seq=2;
			run;

		%end;%else %do;
			data _rpt_;
				set _rpt_ end=stop;
				output;
				if stop then do;
					memname=%upcase("%scan(%scan(&others,&k,%str( )),1,.)");
					note="按 %scan(%scan(&others,&k,%str( )),2,.) ";
					seq=2;output;
				end;
			run;
		%end;
		%let k=%eval(&k+1);
	%end;

	%let k=1;
	%do %while (%scan(&datacopy,&k,%str( )) ne );
		data oData.%scan(%scan(&datacopy,&k,%str( )),1,.);
			set iData.%scan(%scan(&datacopy,&k,%str( )),1,.);
		run;
		%if ^%sysfunc(exist(_rpt_,data)) %then %do;
			data _rpt_;
				length memname note $32;
				memname=%upcase("%scan(&datacopy,&k,%str( ))");
				note="全复制";
				seq=3;
			run;
			%end;%else %do;
			data _rpt_;
				set _rpt_ end=stop;
				output;
				if stop then do;
					memname=%upcase("%scan(&datacopy,&k,%str( ))");
					note="全复制";
					seq=3;output;
				end;
			run;
		%end;

		%let k=%eval(&k+1);
	%end;

	data _tmp_;
		set sashelp.vtable(where=(libname in ("%upcase(iData)", "%upcase(oData)")  and memtype="DATA") keep=libname memtype memname nlobs);
	run;
	proc sort data=_tmp_; by memname libname; run;

	data _tmp_;
		merge _tmp_(where=(libname1="%upcase(iData)") keep=libname memname nlobs rename=(libname=libname1))
		      _tmp_(where=(libname="%upcase(oData)") keep=libname memname nlobs rename=(nlobs=nlobs1));
		by memname;
	run;
	proc sort data=_rpt_; by memname seq; run;
	data _rpt_(drop=seq);
		set _rpt_;
		by memname seq;
		if last.memname;
	run;
	proc sort data=_tmp_; by memname; run;

	data _rpt_(drop=libname libname1);
		merge _tmp_ _rpt_;
		if nlobs>0 then pct=put(nlobs1*100/nlobs,8.1)||"%";
		if nlobs=0 then pct="     --";
		by memname;
		label memname='数据集' nlobs='原记录数' nlobs1='新记录数' note ='提取方法' pct='提取百分率';
	run;
	proc sql; drop table _tmp_;quit;

	proc sort data=_rpt_ out=oData._rpt_; by memname nlobs1 ; run;
	proc sort data=_rpt_; by nlobs1 memname; run;

%mend;


