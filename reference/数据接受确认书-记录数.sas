%macro JILU(Base=,word=);
libname raw "&Base";
proc contents data=raw._all_ out=alldata noprint;
run;proc sort data=alldata;by MEMNAME VARNUM;run;
data final;
set alldata;TYPEMEMaaaa="&word";
by MEMNAME VARNUM;
if last.MEMNAME;
keep MEMNAME VARNUM NOBS TYPEMEMaaaa;
run;
data jilu;set final;e=cats(MEMNAME,"行",NOBS,",","列",VARNUM);drop  VARNUM NOBS;label e='结果';run;
proc sql;
	drop table alldata,final;
quit;
%mend;
%jilu(Base=\\tsclient\Z\DM\金赛031扩展期\e-TMF模板CN-V7.1\29.其他文档\12rawdata,word=031扩展期-Rawdata-20220111)
