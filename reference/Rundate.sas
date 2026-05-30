/**************************************************************
* Program:      Macro used for generate the running date
* Function:     Update running date to today's date when the observation updated;
* Date:         Dec 05, 2018
* Programmer:   Weiyu Wang
*
* keyvar: The key variables used to locate observations, separated by space;
* exvar: Define variables that do not need to be compared, separated by space;
*
* Example: 
%Rundate(oldlib=old,oldset=_mh2,newlib=work,newset=_mh2,outlib=work,outset=rst,keyvar=usubjid PAGE LINE,exvar=STUDYID SITE FORM rundate);
*
* All Right Reserved 
*
* Change history:
*
*
***************************************************************/
%macro Rundate(oldlib=,oldset=,newlib=,newset=,outlib=,outset=,keyvar=,exvar=);

data _null_;
	call symput("exvars",""""||tranwrd("&keyvar. &exvar.",' ','" "')||"""");
run;
%put &exvars;

proc sql noprint;
	select name into:vars separated by " " from SAShelp.vcolumn
	where libname="%upcase(&newlib)" and memname="%upcase(&newset)" and memtype="DATA" and upcase(name) ^in ( %upcase(&exvars) );
quit;

data _null_;
	call symput("cvarnum",put(count("&vars",' ')+1,8.));
run;
proc sort data=&newlib..&newset.;by &keyvar.;run;
proc sort data=&oldlib..&oldset.;by &keyvar.;run;


data &outlib..&outset.;
	merge 	
			&oldlib..&oldset.(rename=(%do ii=1 %to &cvarnum. %by 1; %scan(&vars.,&ii,%str( ))=_%scan(&vars.,&ii,%str( )) %end; rundate=_rundate) in=z98)
&newlib..&newset.(in=z99);
	by &keyvar;
	if z99;
	if z98 %do ii=1 %to &cvarnum. %by 1;and %scan(&vars.,&ii,%str( ))=_%scan(&vars.,&ii,%str( )) %end; then rundate=_rundate;
	else rundate="&sysdate9.";
	drop %do ii=1 %to &cvarnum. %by 1;_%scan(&vars.,&ii,%str( )) %end; _rundate;
run;

%mend Rundate;
