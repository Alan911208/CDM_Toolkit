/**************************************************************
* Program:      SAS2CSV
* Function:     SAS Datasets 2 CSV File Separately
* Date:         Apr 7, 2017
* Programmer:   Jiajun Shen
*
* Clinflash Applications for Clinical Data Management
* 
*
*
* Sample Usage:
  %SAS2CSV(inpath=D:\Project\SCT400NHL2\12 DERIVED,outpath=D:\Project\SCT400NHL2\csv);
  %SAS2CSV(inpath=D:\Project\SCT400NHL2\12 DERIVED,outpath=);
  %SAS2CSV(inpath=D:\Project\SCT400NHL2\12 DERIVED);
*
*
* All Right Reserved 
*
* Change history:
*
*
***************************************************************/

%Macro SAS2CSV(inpath=,outpath=);

    %if %length(&outpath)=0 %then %let outpath=&inpath.;

    libname inlib "&inpath.";
    options nofmterr fmtsearch=(inlib.formats) notes;

    proc sql noprint;
         select memname into:insets separated by " " from sashelp.vtable
         where libname="%upcase(inlib)"  ;
    quit;

    %put &insets;

    %let i=1;
	     %do %while (%scan(&insets,&i,%str( )) ne  );

             PROC EXPORT DATA=inlib.%scan(&insets,&i,%str( )) OUTFILE= "&outpath.\%scan(&insets,&i,%str( )).csv" 
                         DBMS=CSV REPLACE LABEL;
             RUN;

	         %let i=%eval(&i+1);
	     %end;
%mend;

