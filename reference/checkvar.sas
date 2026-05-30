/**************************************************************
* Program:      checkvar
* Function:    判断数据集中是否存在该变量;
				输出一个宏变量&rc
			   *&rc=0    不存在;
			   *&rc=1    存在;

				dsn=数据集名;
				var=变量名;

* Date:         JAN 13,2014
*
*
* Change history:
*
*
***************************************************************/

%macro checkvar(dsn=, var=);
   ***Note: 0---not found :: 1---found;
    proc contents data=&dsn out=_name_(keep=name) noprint; run;
     %global rc;
     %let rc=0;

    data _null_;
      set _name_;
      call symputx('rc','1');
      where upcase(name)="%upcase(&var)";
    run;
	proc sql;
	drop table _name_;
	quit;
    %put rc=&rc  ;

%mend;
