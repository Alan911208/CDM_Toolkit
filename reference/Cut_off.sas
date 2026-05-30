/**************************************************************
* Program:      Cut_off
* Function:     将一个长度超过200的变量截成多个长度为200的变量;
* Date:         Dec 25, 2016
* Programmer:   
*
* Clinflash Applications for Clinical Data Management
* Sample Usage:
* 
*
* All Right Reserved 
*
* Change history:
*
***************************************************************/

%macro cut_off(inset=,var=);

data cut1;
    set &inset.;
	n=ceilz(lengthn(ktrim(&var.))/200);
run;

proc sql noprint;
    select max(n) into:max from cut1;
quit;

%put 需分割成的最大变量个数：&max;

data _null_;
    set cut1;
	call symputx("label",vlabel(&var.));
run;

%put 被分割的变量名和标签：%upcase(&var) &label;

/*data &inset.;
    set cut1;
	tot=1;
	array aa $200 &var.1-&var.%trim(&max);
	%do i=1 %to &max;
	     klth&i.=klength(substr(kleft(&var.),200*%eval(&i.-1)+1,200));
	     &&var.&i=strip(ksubstr(kleft(&var.),tot,klth&i.));
		 tot=tot+klth&i.;
		 label &&var.&i="&&label.&i";
	%end;
	drop n &var. tot klth:;
run;*/

data &inset.;
    set cut1;
	tot=1;
	rtot=1;
	array aa $200. &var.1-&var.%trim(&max);
	%do  i=1  %to  &max;
	     if ktrim(substr(kleft(&var.),rtot,200)) = ktrim(substr(kleft(&var.),rtot,199)) then rlth&i.=length(substr(kleft(&var.),rtot,199));
            else rlth&i.=length(substr(kleft(&var.),rtot,200));

	     klth&i.=klength( substr(kleft(&var.),rtot, rlth&i.));

	     if klth&i.>0 then &&var.&i.=strip(ksubstr(kleft(&var.),tot, klth&i.));

		  tot= tot+klth&i.;
		 rtot=rtot+rlth&i.;
		 label &&var.&i.="&&label.&i.";
	%end;
	drop n &var. tot rtot rlth: klth:;
run;


proc sql noprint;
    drop table cut1;
quit;

%mend cut_off;
