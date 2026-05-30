/**************************************************************
* Program:      cutvara
* Function:     批量删除同后缀的变量;
* Date:         20221207	
* Programmer:   luojiahao
**
**
%cutvara(r1=C:\Users\jiahao.luo\Desktop\xuanshi\正式数据,r2=C:\Users\jiahao.luo\Desktop\xuanshi\nnn);
复制上方语句，r1为输入路径，r2为输出路径，后缀默认为_u如需修改可修改宏第七行备注处，如整个路径下数据集都没有改后缀的变量，会报错
改程序运行会有大量绿色报错警告，可以不管
* Change history:
***************************************************************/
%macro cutvara(r1=,r2=);

libname raw1 "&r1.";
libname raw2 "&r2.";
proc sql;
        select name into: keeplist separated by ' '
        from sashelp.vcolumn
        where libname='RAW1 '  and name like '%_u';
/*		可以修改%后内容更改删除的后缀*/
quit;
/*data raw2.CM;*/
/*  set raw1.CM(drop=&keeplist);*/
/*run;*/
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
	drop &keeplist;
run;
%let kk=%eval(&kk+1);
%end;
%mend;
