/**************************************************************
* Program:      DMR_Q1
* Function:     Generate Query Summary for Data Management Report 
* Result:       T51, T52, T53
* Date:         Sep 10, 2017
* Programmer:   Fangqiu Xiao
*
*
* Apply to Generate Query Summary for Data Management Report from Clinflash CDM System
* Clinflash Applications for Clinical Data Management
*
* Sample Usage:
* %DMR_Qnew(qSheet=query,Form=Form,fDMR=D:\work\xiaofangqiu\DMR_Query.doc，ver=3.0);
*
* Parameters:
* qSheet=Excel Sheet Name of QUERY 
* fDMR=Microsoft Word File
* ver=周报版本
把SDS中的Form表单复制在query表单所在的EXCEL
* All Right Reserved 
*
* Change history:
* 增加表5.2.1 各中心质疑情况表

*根据最新质疑明细调整了列和疑问管理中的主要问题<间接报告>日期计算方式 （luojiahao 2021-12-01)
*使用SDS表单，修复了之前用字段编号可能导致的一些问题（luojiahao 2023-02-02)
*增加版本默认版本/3.0版本Weijian Zhang  2025-09-08)
***************************************************************/
%macro DMR_Qnew(qSheet=, Form=, fDMR=, ver=);
    %xls2sas(insheet=&qSheet, outset=query, row=50000);
    %xls2sas(insheet=&Form, outset=q1, row=50000);

    %if %upcase(&ver) = DEFAULT %then %do;
        data q2;
            set q1;
            keep _c1_ _c2_;
            rename _c1_=c _c2_=_c8_;
        run;

        proc sort data=q2; by _c8_; run;
        proc sort data=query; by _c8_; run;
        data query;
            merge query q2;
            by _c8_;
            if _c1_ = '' then delete;
        run;

        data query;
            set query(where=(_c21_ ^in ("取消","取","Cancel")));
            rename _c3_=site_name
                   _c4_=subject_No
                   _c8_=form_name
                   c=Field_OID
                   _c13_=query_name
                   _c14_=Query_Open_Date
                   _c17_=query_text
                   _c18_=Answer_Date
                   _c21_=Query_Status
                   _c22_=Resolve_Date;
        run;
    %end;
    %else %if %upcase(&ver) = 3.0 %then %do;
        data q2;
            set q1;
            keep _c1_ _c2_;
            rename _c1_=c _c2_=_c9_;
        run;

        proc sort data=q2; by _c9_; run;
        proc sort data=query; by _c9_; run;
        data query;
            merge query q2;
            by _c9_;
            if _c2_ = '' then delete;
        run;

        data query;
            set query(where=(_c22_ ^in ("取消","取","Cancel")));
            rename _c4_=site_name
                   _c5_=subject_No
                   _c9_=form_name
                   c=Field_OID
                   _c14_=query_name
                   _c15_=Query_Open_Date
                   _c18_=query_text
                   _c19_=Answer_Date
                   _c22_=Query_Status
                   _c23_=Resolve_Date;
        run;
    %end;
    %else %do;
        %put ERROR: Unsupported version: &ver;
        %return;
    %end;

    /* 后续逻辑不变 */
    proc sql noprint;
        select count(*), count(DISTINCT site_name), count(DISTINCT subject_No)
        into :nTotal, :nSite, :nSubj
        from query (where=(Query_Status ^in ("取消","取","Cancel")));
    quit;

    

	data t51;
		length formid $100;
		set query (where =(Query_Status ^in ("取消","取","Cancel")));
		formid="1"||substr(Field_OID,1,2);output;
		formid="1"||substr(Field_OID,1,2)||":"||left(form_name);output;
		formid="9";output;
	run;

	proc sort data= t51; by formid  query_text; run;

	data  t51(keep=seq formid form_name query_text nqry);
		retain seq 1.1 nFrmID nTxt 0;
		set t51;by formid query_text;
		nFrmID=nFrmId+1; nTxt= nTxt +1;
		if last.query_text then do;query_text=strip(query_text)||compress("["||nTxt||"]"); nQry=nTxt;output; nTxt=0; end;
		if last.formid then do;query_text=nFrmId||"";seq=seq-0.1; nQry=nFrmId ;output;seq=seq+1.1;nFrmID=0;end;
	run;

	proc sort data= t51; by formid seq descending nqry; run;
	data t51;
		retain _line_ 0;
		set t51;by formid seq descending nqry; 
		if seq-int(seq) lt 0.09 then  _line_=0;
		nseq=seq+_line_/1000;
		_line_=_line_+1;
		if _line_<5 then output;
	run;

	proc transpose data=t51 out=t51(keep=formid line1-line4)  prefix=line;
		var query_text;
		by formid ;
		id _line_;
	run;

	data t51(keep=formid--line);
		length formid $100 line1 $200 line $1000;
		set t51;
		if ^missing(line2) then line="(1) "||strip(line2);
		if ^missing(line3) then line=strip(line)||"; (2) "||strip(line3);
		if ^missing(line4) then line=strip(line)||"; (3) "||strip(line4);
	run;

	data t51;
		set t51 end=stop;
		if substr(formid,4,1)=":" then formid=" "||substr(formid,5);else formid=substr(formid,2);
		if Missing(formid) then formid="Total"; output;
		if stop then do;
			formid="平均（/受试者）";line1=int(&nTotal*10/&nSubj)/10;line='';output;
			formid="平均（/研究机构）";line1=int(&nTotal*10/&nSite)/10;output;
		end;
	    label formid='数据集/数据块' line1='质疑数' line='<<前三位质疑文本>>';
	run;

	title1 5.1疑问的总体情况<间接报告>;
	footnote ;
	%mReport(inset=t51,append=new,filetype=rtf,file=&fDMR);

	data oQry(keep=site_name form_name Query_Name query_text Resolve_Date Query_Open_Date qDate);
		set Query;
		if missing(Answer_Date) then Answer_Date=Resolve_Date;
		qDate=input(substr(Answer_Date,1,10),yymmdd10.)-input(substr(Query_Open_Date,1,10),yymmdd10.);
	run;

	proc means data=oQry n mean min max noprint;
		var qDate;
		class Query_Name;
		output out=t52;
	run;

	proc transpose data=t52 out=t52;
		var qdate;
		by Query_Name;
		id _stat_;
	run;

	data t52(keep=type omean orange rmean rrange);
		length type omean orange rmean rrange $50; 
		set t52;
	 	type=strip(Query_Name)||compress("("||n||")");
		if Query_Name='' then type='总体'||compress("("||n||")");
		omean='无信息';orange='无信息';
		rmean=putn(mean,8.1);
		rrange=compress(min||"-"||max);
		type=" "||type;
		label type='质疑类型' omean='平均天数(产生)' orange='范围(产生)' rmean='平均天数(回复)' rrange='范围(回复)';
	run;

	title1 5.2 疑问的处理情况<间接报告>;
	footnote 注意：Query报表不包含疑问产生信息,疑问回复时间=质疑回答日-质疑产生日，质疑直接解决的,质疑回答日为质疑解决日;
	%let header1=!!疑问产生天数^!疑问回复^!;
	%let header2=!质疑类型!平均天数 (天)!范围 （天）!平均天数 (天)!范围 （天）!;
	%mReport(inset=t52,headrow=2,append=yes,filetype=rtf,file=&fDMR);

	*增加表5.2.1 各中心质疑分类情况;
	proc means data=oQry n mean median min max noprint;
		var qDate;
		class site_name;
		output out=t521 N=N MEAN=MEAN MEDIAN=MEDIAN MIN=MIN MAX=MAX;
	run;

	data t521(keep=type nc rmean rmedian rrange);
		length type nc rmean rmedian rrange $50; 
		set t521;
	 	type=strip(site_name);
		if site_name='' then type='总体';
		nc=putn(n,8.);
		rmean=putn(mean,8.1);
		rmedian=putn(median,8.);
		rrange=compress(min||"-"||max);
		type=" "||type;
		label type='中心号' nc='质疑数量' rmean='平均天数' rmedian='中位天数' rrange='范围';
	run;

	proc sort data=t521;by type;run;

	title1 5.2.1 各中心疑问的处理情况<间接报告>;
	footnote;
	%let header1=!中心号!疑问数量!平均天数（天）!中位天数（天）!范围（天）!;
	%let header2=!（列出中心号）!列出质疑总数!（列出各中心疑问生成到答疑的平均天数）!（列出各中心疑问生成到答疑的中位天数）!（列出最短天数~最长天数范围）!;
	%mReport(inset=t521,headrow=2,append=yes,filetype=rtf,file=&fDMR);

	/******************************************/

	proc sort data=oQry; by site_name; run;

	data oQry1(keep=site_name nqry);
		retain n0 n1 t0 0;
		set oQry; by site_name;
		if first.site_name then n0=0;
		n0=n0+1;
		if last.site_name then do; nQry=n0; output;end;
	run;
	proc sort data=oQry1; by descending nqry; run;
	data oQry1(drop=n0);
		retain n0 0;
		set oQry1;by descending nqry;
		if _n_=1 then n0=nqry;
		if nqry=n0;
	run;

	

	proc sort data=oQry; by descending qdate; run;
	data oQry2(keep=site_name qdate form_name Query_Name query_text);
		retain n0 0;
		set oQry;by descending qdate;
		if _n_=1 then n0=qdate;
		if qdate=n0;
	run;
data oQry2;set oQry2;qdate=qdate-1;run;
	data T53(keep=qtype qsite qrpt qcomment);
		length qtype qsite qrpt qcomment $100;
		set oqry2(in=a) oqry1(in=b);
		if a then do;
			qtype=' 疑问回复时间最长研究单位及原因分析';
			qsite=site_name;
			qrpt=compress(qdate)||"(天)" ;
			qcomment=strip(form_name)||"/"||strip(query_text)||"/"||strip(Query_Name);
		end;
		if b then do;
			qtype=' 疑问产生数量最高研究单位及原因分析';
			qsite=site_name;
			qrpt=compress(nqry||"("||put(nqry*100/&nTotal,8.1)||"%)") ;
			qcomment='<<填写原因及说明>>';
		end;
		label Qtype='疑问重点问题' qsite='研究单位名称' qrpt='时间/数量' qcomment='原因及说明';
	run;

	title 5.3疑问管理中的主要问题<间接报告>;
	footnote 注意：疑问回复时间=质疑回答日-质疑产生日.质疑直接解决的,质疑回答日为质疑解决日;
	%mReport(inset=t53,append=yes,filetype=rtf,file=&fDMR);
	proc sql;
		drop table oQry,oQry1,oQry2;
	quit;
%mend;



/* 3.0版本 
%DMR_Qnew(qSheet=QueryDetailReport, Form=Form, fDMR=C:\Users\weijian.zhang\Desktop\DMR_Query.doc, ver=3.0);*/
/* 默认版本 
%DMR_Qnew(qSheet=QueryDetailReport, Form=Form, fDMR=C:\Users\weijian.zhang\Desktop\DMR_Query.doc, ver=default); */
