/**************************************************************
* Program:      DelEmpty
* Function:     to Remove duplicated blank observations;
* Date:         Mar 20, 2008
* Programmer:   Fangqiu Xiao
*
* Apply to check data extracted from ClinStReport CDM System
* ClinStReport Applications for Clinical Data Management
*
* Sample Usage:
* %delEmpty(insets=derived.ae derive.cm);
* %delEmpty(insets=derived.ae derive.cm,outsets=derived.ae derive.cm);
*
*
* DropEXTR=yes: NOT include blank Extra pages;
* rptDrop=YES: to remove duplicate blank rows such as in AE, CM ect.
* All Right Reserved 
*
* Change history:
*
*
***************************************************************/

%macro DelEmpty(inlib=,outlib=,insets=,outsets=,vIndex=SEQNUM,pstatus=NO);
	%local exData;
	%let exData="WVISIT","LABREF","CMEDDRA","CUSERD","MISPAGE","CWHODD","PAGESTATUS","CROSSCHK","TRANSLATION","QUERY","FORMATS","PAT";
	%if %length(&inlib)=0 %then %let inlib=work;
	%if %length(&outlib)=0 %then %let outlib=&inlib;
	%if %length(&insets)=0 %then %do;
		proc sql noprint;
			/*select memname into:insets separated by " " from sashelp.vtable */
			/*where libname="%upcase(&inlib)" and memtype="DATA" and memname ^in ( &exData );*/
			select memname into:insets separated by " " from sashelp.vcolumn
			where upcase(libname)=upcase("derived") and  memtype="DATA" and name="pStatus";		quit;
	%end;
	%if %length(&outsets)=0 %then %let outsets=&insets;

	%let i=1;
	proc sql noprint;
		create table _tmp_
		as select * from sashelp.vcolumn
		where upcase(libname)=upcase("&Inlib") and  memtype="DATA";
	quit;

	%do %while (%scan(&insets,&i,%str( )) ne  );
		%let inset=%scan(&insets,&i,%str( ));
		%let outset=%scan(&outsets,&i,%str( ));;
		%let estINX=0; /* Not existed  */  %let misINX=0; /* All missing */
		%let estCD=0; /* Not existed  */   %let misCD=0; /* All missing */
		%let estTEXT=0; /* Not existed  */ %let misTEXT=0; /* All missing */
		proc sql noprint;
			select count(name) into : estINX from _tmp_
			where upcase(libname)=upcase("&Inlib") and upcase(memname)=upcase("&inset")
			and upcase(name)=upcase("&vIndex") and memtype="DATA";
			select count(name) into : estCD from _tmp_
			where upcase(libname)=upcase("&Inlib") and upcase(memname)=upcase("&inset")
			and upcase(name)=upcase("seqCD") and memtype="DATA";
			select count(name) into : estTEXT from _tmp_
			where upcase(libname)=upcase("&Inlib") and upcase(memname)=upcase("&inset")
			and upcase(name)=upcase("seqTEXT") and memtype="DATA";
			/* add on 2008-11-04 */
			select format into : INXfmt from _tmp_
			where upcase(libname)=upcase("&Inlib") and upcase(memname)=upcase("&inset")
			and upcase(name)=upcase("&vIndex") and memtype="DATA";
			/* end of add */
		quit;

		%if &estINX %then %do;
			proc sql noprint;
				select count(&vIndex) into : misINX  from &inlib..&inset where ^missing(&vIndex);
			quit;
		%end;
		%if &estCD %then %do;
			proc sql noprint;
				select count(seqCD) into : misCD  from &inlib..&inset where ^missing(seqCD);
			quit;
		%end;
		%if &estTEXT %then %do;
			proc sql noprint;
				select count(seqTEXT) into : misTEXT  from &inlib..&inset where ^missing(seqTEXT);
			quit;
		%end;
		data &outlib..&outset;
			set &inlib..&inset(where=(^missing(line)));
			%if &estINX and (&misINX=0) %then %do;
				if (Status in (.,0) and line ne 1) or (pStatus='Added    ') then delete;
				%if &estINX and %sysfunc(compress(&INXfmt,0123456789.)) eq  %then drop &vIndex;;
			%end;;
			%if &estINX  %then %do;
				%if %sysfunc(compress(&INXfmt,0123456789.)) ne  %then if missing(&vIndex) then &vIndex=line;;
			%end;;
			%if &estCD and (&misCD=0)  %then drop seqCD;;
			%if &estTEXT and (&misTEXT=0)  %then drop seqTEXT;;
			%if %upcase(&pStatus) ne YES   %then drop pStatus;;

		run;
		%let i=%eval(&i+1);
	%end;
	proc sql noprint;
		drop table _tmp_;
	quit;
	%if %sysfunc(exist(&outlib..pat,data)) %then %do;
		proc sql; drop table &outlib..pat; quit;
	%end;

%mend;

