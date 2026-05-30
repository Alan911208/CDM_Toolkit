 /**************************************************************
* Program:      export
* Function:     Convert SAS Data Sets to Excel Sheets in batch Mode;
* Date:         Dec 06,2022
* Programmer:   Wenkai Hua
*
* Clinflash Applications for Clinical Data Management
* 
* Sample Usage:

* %export(libname,address,excelname,xlsx,xlsx);
* %export(out,D:\ClinStReport,test,xlsx,xlsx);

* All Right Reserved 
*
* Change history:
*
***************************************************************/


%macro export(libname,outfile,filename,type,dbms);
	%let libname=%upcase(&libname);
	proc sql noprint;
	select memname,count(memname) into : memlist separated by '\', : nummem
	from dictionary.tables
	where libname="&libname";
	run;
	quit;

	%do i=1 %to &nummem;
		%let memname=%scan(&memlist,&i,\);
		proc export data=&libname..&memname
		outfile="&outfile\&filename..&type"
		dbms=&dbms replace label;
		sheet="&memname";
	run;
		%end;
%mend export;
