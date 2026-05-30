**************************************************************************************************;
*Output to Excel(compare.xlsx);
*Add Colour;

/*Usage
0-run origin date

 %librundateo(pLib=C:\Users\jiahao.ma\Desktop\附件\MM listing程序1上一次的out,rundate='20220201');

1-run the macro

2-*Update dataset path, do not change Libname
  *Update excvar（no need comparison variable) &DropVar & Keyval（Key data point）
  *Run the six line program below

libname data"C:\Users\jiahao.ma\Desktop\附件\out";
libname last"C:\Users\jiahao.ma\Desktop\附件\MM listing程序1上一次的out";
libname out_f"C:\Users\jiahao.ma\Desktop\1";

%let ExcVar="_STATUE","RUNDATE","USUBJID", "SITEID", "VISIT", "PAGE", "FORM", "LINE","STATUS";
%let DropVar=SDVTier folderid;
%runexportandcolour(inLib=data,OldLib=last,Outlib=out_f,keyvar=USUBJID SITEID VISIT PAGE FORM LINE);

/
3-*Please open the compare.xlsx first and run the program below.


%runaddcolor(OutLib=out_f);

*/
**************************************************************************************************;
%macro runexportandcolour(inLib= ,OldLib=,OutLib=,keyvar=);


	data tmp;
		set sashelp.vtable(where=(libname='DATA'));
	run;

	proc sql noprint;
		select memname into:name separated by ' ' from tmp;
	quit;
	%let nowdate = %sysfunc(date(), yymmddn8.);
	%let dsetnum=1;
	%do %while (%scan(&name, &dsetnum) ne );
	%let dset=%scan(&name, &dsetnum);


	

		data &OutLib..&dset.;
			set &inlib..&dset.;
			drop &DropVar;
		run;
		data &dset.;
			set &OldLib..&dset.;
			drop &DropVar;
			if _STATUE='Inact' then delete;
		run;
			proc contents noprint data=&OutLib..&dset. out=Var(keep=MEMNAME VARNUM NAME where=(upcase(NAME) ^in (&ExcVar)));run; 
			proc contents noprint data=&dset. out=OldVar(keep=MEMNAME VARNUM NAME where=(upcase(NAME) ^in (&ExcVar)));run; 

			proc sort data=Var;by MEMNAME NAME;run;
			proc sort data=OldVar;by MEMNAME NAME;run;

			proc sort data=&OutLib..&dset  out=&OutLib..&dset;by &keyvar;run;
			proc sort data=&dset  out=&dset;by &keyvar;run;

			data T2_&dset;
				merge &dset(in=a)
					  &OutLib..&dset(in=b);
				by &keyvar;
				if a and ^b;
				_Statue="Inact";
				RUNDATE="&nowdate.";
			run;

			data _null_;
				merge Var(in=a) OldVar(in=b);
				by MEMNAME NAME;
				if a;
				length RenameVar DIFFVar NewVar $10000;
				retain RenameVar DIFFVar NewVar;
				call symput('Col'||left(put(_n_,3.)),compress(VARNUM));*get the sequence of var ;
				call symput('VarName'||left(put(_n_,3.)),compress(NAME));*get the id of var;
				if first.MEMNAME then do;RenameVar="";DIFFVar="";NewVar=" ";end;
				if a and b then RenameVar=strip(strip(RenameVar)||" "||strip(NAME)||"=last_"||strip(NAME));
				if a and ^b then NewVar=strip(strip(NewVar)||"last_"||strip(NAME))||"=''; ";

				DIFFVar=strip(strip(DIFFVar)||" DIFF"||strip(VARNUM));
				if last.MEMNAME then do;
					Call Symput("RenameVar",RenameVar);
					Call Symput("NewVar",NewVar);
					Call Symput("DIFFVar",DIFFVar);
					call symput('number',left(compress(put(_n_,best.))));
				end;
			run;
			%put RenameVar=&RenameVar;
			%put NewVar=&NewVar;

			%put DIFFVar=&DIFFVar;
			%put number=&number;

			proc sort data=&dset out=&OldLib._&dset(rename=(&RenameVar));by &keyvar;run;
			data &OldLib._&dset;
				set &OldLib._&dset;
				&NewVar;
			run;

			proc sort data=&OutLib..&dset  out=&inLib._&dset;by &keyvar;run;

			data T_&dset(keep=&keyvar Type) 
				 C_&dset(keep=&keyvar &DIFFVar);
				merge &OldLib._&dset(in=a) 
					  &InLib._&dset(in=b);
				by &keyvar;
				length Type &DIFFVar $20;
				if b then do;
					%do i = 1 %to &number.;
						DIFF&&Col&i..="";
						if (upcase(&&VarName&i..)^=upcase(last_&&VarName&i..) & a) then do;DIFF&&Col&i..=compress("C"||&&Col&i..);Type="Update";end;
					%end;
					if ^a then Type="New";	
					output; 
				end;	
			run;
						 
			data T1_&dset(drop=last_RUNDATE Type);
				merge &inLib._&dset(in=a)
					  &OldLib._&dset(keep=&keyvar RUNDATE rename=(RUNDATE=last_RUNDATE))
					  T_&dset;
				by &keyvar;
				if a;
				_Statue=Type;
				if Type ^in ("New","Update") then RUNDATE=last_RUNDATE;
				if RUNDATE='' then RUNDATE="&nowdate.";
			run;


			data &OutLib..&dset.;
				set T1_&dset 
					T2_&dset;
				
			run;

			data freq_&dset;
				set &OutLib..&dset.;
/*				if _Statue^='';*/
			run;

			proc sort data=freq_&dset out=freq_&dset;by &keyvar _Statue;run;

			PROC FREQ DATA=freq_&dset noprint;
			    TABLES _Statue / OUT=freq;
			RUN;

			proc sort data=freq_&dset out=freq1_&dset nodupkey;by usubjid;run;

			data freq1_&dset;
			length A $20;
				set freq1_&dset;
				if usubjid^='' then A="Subject Total";
			run;

			PROC FREQ DATA=freq1_&dset noprint;
			    TABLES A / OUT=freq1;
			RUN;



/*			PROC DATASETS LIB=work nolist;*/
/*    			DELETE freq1_&dset;*/
/*			QUIT;*/

			data freq1;
			length A $20;
			set freq1;
			rename A=_Statue;
/*			if A^='' then A="Subject Total";*/
			run;

			data freq;
			set freq1 freq;
			run;

			data freq;
				set freq;
				length DatasetName $32.;
				DatasetName = "&dset";
/*				if _Statue^='';*/
			run;

						proc sort data=freq out=freq;by _Statue;run;
						PROC SQL noprint;
						   /* Check if value 'b' exists in column A */
						   SELECT COUNT(*) INTO :cnt FROM freq WHERE _Statue = 'Inact';
						QUIT;

						%IF &cnt = 0 %THEN %DO;
						   /* If value 'b' does not exist in column A, insert a new row */
						   PROC SQL noprint;
						      INSERT INTO freq (_Statue,COUNT,PERCENT,DatasetName) VALUES ('Inact',0,0,"&dset" ) ;
						   QUIT;
						%END;

						PROC SQL noprint;
						   /* Check if value 'b' exists in column A */
						   SELECT COUNT(*) INTO :cnt FROM freq WHERE _Statue = 'Update';
						QUIT;

						%IF &cnt = 0 %THEN %DO;
						   /* If value 'b' does not exist in column A, insert a new row */
						   PROC SQL noprint;
						      INSERT INTO freq (_Statue,COUNT,PERCENT,DatasetName) VALUES ('Update',0,0,"&dset" );
						   QUIT;
						%END;


						PROC SQL noprint;
						   /* Check if value 'b' exists in column A */
						   SELECT COUNT(*) INTO :cnt FROM freq WHERE _Statue = 'New';
						QUIT;

						%IF &cnt = 0 %THEN %DO;
						   /* If value 'b' does not exist in column A, insert a new row */
						   PROC SQL noprint;
						      INSERT INTO freq (_Statue,COUNT,PERCENT,DatasetName) VALUES ('New',0,0,"&dset" );
						   QUIT;
						%END;
						PROC SQL noprint;
						   /* Check if value 'b' exists in column A */
						   SELECT COUNT(*) INTO :cnt FROM freq WHERE _Statue = 'Subject Total';
						QUIT;

						%IF &cnt = 0 %THEN %DO;
						   /* If value 'b' does not exist in column A, insert a new row */
						   PROC SQL noprint;
						      INSERT INTO freq (_Statue,COUNT,PERCENT,DatasetName) VALUES ('Subject Total',0,0,"&dset" );
						   QUIT;
						%END;
						
						PROC SQL noprint;
						   /* Check if value 'b' exists in column A */
						   SELECT COUNT(*) INTO :cnt FROM freq WHERE _Statue = '' ;
						QUIT;

						%IF &cnt = 0 %THEN %DO;
						   /* If value 'b' does not exist in column A, insert a new row */
						   PROC SQL noprint;
						      INSERT INTO freq (_Statue,COUNT,PERCENT,DatasetName) VALUES ('total',0,0,"&dset" );
						   QUIT;
						%END;



			data freq;
				set freq;
				if _Statue='' then _Statue='stable';
			run;


			PROC TRANSPOSE DATA=freq
			    OUT=freq_out(drop=_NAME_)
			    PREFIX=Form_;
			    ID _Statue;
				by DatasetName;
			    VAR COUNT;
			RUN;

			data freq_out(drop=Form_stable);
			retain DatasetName _LABEL_ Form_New Form_Update Form_Inact Form_Total Form_Subject_Total;
				set freq_out;
				if Form_stable='' then Form_stable=0;
				Form_Total=form_Inact+form_Update+form_new+Form_stable;
			run;



			PROC APPEND BASE=SourceSummary DATA=freq_out(drop=_LABEL_); RUN;
			


%let dsetnum=%eval(&dsetnum+1);
%end;
		%macro nowpath;
		        %global fullpath nowpath;   
		        proc sql noprint;
		                select xpath into : fullpath   
		                        from  dictionary.extfiles  
		                        where substr(fileref,4) eq
		                                (select max(substr(fileref,4))   
		                                from dictionary.extfiles
		                                where substr(fileref,1,1) eq "#" and index(xpath,".sas") gt 0  );
		        quit;

		        
		        %let nowpath=%substr(&fullpath,1,%eval(%length(&fullpath)-%length(%scan(&fullpath,-1,\))));
		        %put The Full Path is: &fullpath;
		        %put The Now Path is: &nowpath;
		%mend nowpath;

		%nowpath;


%let libpath = %sysfunc(pathname(&inlib));
%let filename = %sysfunc(scan(&libpath, -2, '\'))\%sysfunc(scan(&libpath, -1, '\'));
%put &filename;

data sourcetype1;
    length sourcetype filename $ 256;
	sourcetype="SASdataset New";
    filename = "&filename";
    output;
run;

%let libpath = %sysfunc(pathname(&OldLib));
%let filename = %sysfunc(scan(&libpath, -2, '\'))\%sysfunc(scan(&libpath, -1, '\'));
%put &filename;

data sourcetype2;
    length sourcetype filename $ 256;
	sourcetype="SASdataset Old";
    filename = "&filename";
    output;
run;

data sourcetype;
    set sourcetype1 sourcetype2;
run;



%let xlsxname=Compare&nowdate.;
%put &xlsxname;

	PROC EXPORT DATA=SourceSummary
		            OUTFILE= "&nowpath.\&xlsxname..xlsx" 
		            DBMS=xlsx  REPLACE  ;
		    		SHEET=SourceSummary; 
	RUN;
	PROC EXPORT DATA=sourcetype
		            OUTFILE= "&nowpath.\&xlsxname..xlsx" 
		            DBMS=xlsx  REPLACE  ;
		    		SHEET=sourcetype; 
	RUN;
PROC DATASETS LIB=work nolist;
    DELETE SourceSummary sourcetype sourcetype1 sourcetype2 freq freq1 freq_out;
QUIT;


	proc sql noprint;
		select memname into:name separated by ' ' from tmp;
	quit;
	%let dsetnum=1;
	%do %while (%scan(&name, &dsetnum) ne );
	%let dset=%scan(&name, &dsetnum);
		

		PROC EXPORT DATA=&OutLib..&dset
		            OUTFILE= "&nowpath.\&xlsxname..xlsx" 
		            DBMS=xlsx  REPLACE label ;
		    		SHEET=&dset; 
		RUN;


	%let dsetnum=%eval(&dsetnum+1);
	%end;


%mend;









%macro runaddcolor(OutLib=);

	data tmpc;
		set sashelp.vtable(where=(libname='OUT_F'));
	run;

	proc sql noprint;
		select memname into:name separated by ' ' from tmpc;
	quit;
	%let dsetnumc=1;
	%do %while (%scan(&name, &dsetnumc) ne );
	%let dset=%scan(&name, &dsetnumc);
	%put &dset;



filename cmds dde 'excel|system';

data _null_;
  file cmds;
  put '[workbook.select("' "&dset" '")]';
run;



filename sas2xl dde 'excel|system';
data _null_;
	file sas2xl;
	
	set C_&dset;
	Array _chr_ _CHAR_;
	do over _chr_;
		if index(upcase(vname(_chr_)),"DIFF") and ^missing(_chr_) then do;
			Range=compress('r'||_n_+1||_chr_||':r'||_n_+1||_chr_);
			put '[select("' range '")]'; 
			put '[patterns(1,,37)]';
		end;
	end;
run;

%let dsetnumc=%eval(&dsetnumc+1);
%end;
%mend;

