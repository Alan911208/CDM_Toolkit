/**************************************************************
* Program:      addvar
* Function:     批量增加变量;
* Date:         20220901	
* Programmer:   luojiahao
**
**
%addvar(form=dm1,dat=rficdtc,kdat=usubjid,Path1=D:\project\TG1506pH4\11 Rawdata,Path2=D:\project\TG1506pH4\11 Rawdata\0901);

复制上方语句，dat是变量用空格分开，可填多个，
form只能填一个 ，kdat为排序匹配变量，一般为usubjid，可根据需求加上visit,pageid,line等
注：拼接的变量编号如在原数据集中有，会出现被覆盖的情况
* Change history:
*
***************************************************************/


%macro addvar(form=,dat=,kdat=,Path1=,Path2=);
libname raw1 "&Path1.";
/*原数据集目录*/
libname raw2 "&Path2.";
/*生成数据集目录*/

data a;
set raw1.&form;keep &kdat &dat;run;
proc sort data=a;by &kdat;run;

data tmp;
	set sashelp.vtable(where=(libname='RAW1'));
run;
proc sql noprint;
	select memname into:name separated by ' ' from tmp;
quit;
%let kk=1;
%do %while (%scan(&name, &kk) ne );
proc sort data=raw1.%scan(&name, &kk);by &kdat;run;

data raw2.%scan(&name, &kk);
	merge raw1.%scan(&name, &kk) a;
if subjid^='';by &kdat;
run;
%let kk=%eval(&kk+1);
%end;
%mend;
