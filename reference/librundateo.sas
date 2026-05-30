/**************************************************************
* Program:      librundateo
* Function:     库中数据集里统一加上rundate列;
* Date:         Sep 22, 2020
* Programmer:   Qiongyu Shen
*
* Clinflash Applications for Clinical Data Management
* 
* Sample Usage:

  %librundateo(pLib=C:\Users\Administrator\Downloads\TG1902EPL_V1_0_20200901_13_46_11,rundate='07JUL2020');

* Change history:
* 
*
***************************************************************/

%macro librundateo(pLib=,rundate=);
libname tmp1 "&pLib";

data tmp;
	set sashelp.vtable(where=(libname='TMP1'));
run;

proc sql noprint;
	select memname into:name separated by ' ' from tmp;
quit;

%let kk=1;
%do %while (%scan(&name, &kk) ne );

data tmp1.%scan(&name, &kk);
	set tmp1.%scan(&name, &kk);
	rundate=&rundate;
run;

%let kk=%eval(&kk+1);
%end;


%mend;


/*%librundateo(pLib=C:\Users\Administrator\Downloads\TG1902EPL_V1_0_20200901_13_46_11,rundate='07JUL2020');*/
