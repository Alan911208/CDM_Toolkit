
/**************************************************************
* Program:      librundatef
* Function:     将库里的数据集批量进行rundate;
* Date:         Sep 22, 2020
* Programmer:   Qiongyu Shen
*
* keyvar: The key variables used to locate observations, separated by space;
* exvar: Define variables that do not need to be compared, separated by space;
*
*
* Sample Usage:

  %librundatef(
	pNew=D:\temp\tmp\old,
	pOld=D:\temp\tmp\new,
    pFin=D:\temp\tmp\fin,
	keyvar=usubjid PAGE LINE,
    exvar=STUDYID SITE FORM rundate);

* Change history:
* 
*
***************************************************************/

%macro librundatef(pNew=,pOld=,pFin=,keyvar=,exvar=);
libname new "&pNew";
libname tmp1 "&pOld";
libname fin "&pFin";
 
data tmp;
	set sashelp.vtable(where=(libname='TMP1'));
run;

proc sql noprint;
	select memname into:name separated by ' ' from tmp;
quit;

%let kk=1;
%do %while (%scan(&name, &kk) ne );

%let dd=%scan(&name, &kk);
%put &dd;

%Rundate(oldlib=tmp1,oldset=&dd,newlib=new,newset=&dd,outlib=fin,outset=&dd,keyvar=&keyvar,exvar=&exvar);

%let kk=%eval(&kk+1);
%end;
%mend;


/*%librundatef(pNew=C:\Users\Administrator\Desktop\new,*/
/*pOld=C:\Users\Administrator\Desktop\old,*/
/*pFin=C:\Users\Administrator\Desktop\fin,*/
/*keyvar=usubjid PAGE LINE,exvar=STUDYID SITE FORM rundate);*/
