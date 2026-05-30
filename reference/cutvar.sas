/**************************************************************
* Program:      cutvar
* Function:     批量删除变量;
* Date:         20220901	
* Programmer:   luojiahao
**
**
%cutvar(dat=__STUDYOID,Path1=D:\project\TG1506pH4\11 Rawdata,Path2=D:\project\TG1506pH4\11 Rawdata\0901);
复制上方语句，变量用空格分开
*
* Change history:
*
***************************************************************/


%macro cutvar(dat=,Path1=,Path2=);
libname raw1 "&Path1.";
/*原数据集目录*/
libname raw2 "&Path2.";
/*生成数据集目录*/
data tmp;
	set sashelp.vtable(where=(libname='RAW1'));
run;
proc sql noprint;
	select memname into:name separated by ' ' from tmp;
quit;
%let kk=1;
%do %while (%scan(&name, &kk) ne );
data raw2.%scan(&name, &kk);
	set raw1.%scan(&name, &kk);
drop &dat;
run;
%let kk=%eval(&kk+1);
%end;
%mend;
