libname DateVar "C:\Users\jiahao.ma\Desktop\old";

%macro getdate();

    %if %sysfunc(exist(datelist,data)) %then %do;
       proc sql noprint;
           drop table datelist;
       quit;
    %end;

    proc sql noprint;
       select memname into :inList separated by ' ' 
       from sashelp.vtable where libname="DATEVAR" and memtype="DATA" and (upcase(memname) ne "FORMATS") ;
    quit;

    %let k=1;
    %do %while (%scan(&inList,&k,%str( )) ne );
       data _date(keep=studyid--line vName vLabel value);
           set DateVar.%scan(&inList,&k,%str( ));
           length vName $20 vLabel $200 value $200;
           array s (*) _CHAR_;
           do i=1 to dim(s);
              do j=1 to dim (s);
                  if (index(upcase(vname(s(i))),'DTC') or index(upcase(vname(s(i))),'DAT') or index(upcase(vname(s(i))),'TIM')) & ^index(upcase(vname(s(i))),'_') then do;
                     vName=vname(s(i));
                     vLabel=vlabel(s(i));
                     value=s(i);
                     if ^missing(value) then output;
                  end;
              end;
           end;
       run;

       proc sort data=_date out=_date nodupkey; by studyid--value; run;

       data datelist;
           set %if %sysfunc(exist(datelist,data)) %then datelist; _date;
       run;
       proc sql noprint;
           drop table _date;
       quit;
       %let k=%eval(&k+1);
    %end;

%mend;
%getdate;
data DateVar.datelist;
set datelist;
run;
