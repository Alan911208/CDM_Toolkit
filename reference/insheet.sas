/**************************************************************
* Program:      Insheet
* Function:     EXCEL中非中文SHEET批量导入SAS;
* Date:         2022-8-1
* Programmer:   Luojiahao

* Sample Usage:

%insheet(raw=C:\Users\jiahao.luo\Desktop\MMLIST20220729.xlsx,raw1=C:\Users\jiahao.luo\Desktop\新建);


raw=一般为XLS,XLSX格式	文件名不要有特殊字符


***************************************************************/

%macro insheet(raw=,raw1=);

libname raw "&raw";
libname raw1 "&raw1";
proc contents data=raw._ALL_ memtype=data out=excel(keep=memname) noprint;run;
proc sort data=excel;by memname;run;
data excel1;set excel;by memname;
if first.memname;
sheet=memname;format sheet $32.;
 keep sheet;run;
data excel;set excel1;sheet1=compress(sheet,'$');
run;
libname raw clear;

%macro imp;
proc sql noprint;
select sheet1,sheet,count(*)
into :sheetlist separated by ' ',
:sheet1list separated by ' ',
     :num
  from excel;
  quit;
  %put sheet=&sheetlist;
  %put sheet1=&sheet1list;
  %put num=&num;

  %do i=1 %to &num;
  %let sheet=%scan(&sheetlist,&i,%str( ));
  %let sheet1=%scan(&sheet1list,&i,%str( ));

proc import datafile="&raw" out=raw1.&sheet dbms=excel  replace;
sheet="&sheet"; GETNAMES=YES;
run;

   %end;

  %mend imp;
  %imp;
%mend;


