/**************************************************************   
* Clinflash Applications for Clinical Data Management                                                                                   
* Program:    肿瘤疗效，一致性核查程序                                                                                      
* Function:     跑肿瘤疗效和一致性核对                                                                                                     
* Date:         2024.6.18                                                                                                         
* Programmer:   luojiahao                                                                                                                   
*                                                                                                                                       
* Clinflash Applications for Clinical Data Management                                                                                                                                         
                                                                                                            
*Instruction：                                                                                                                          
%RS(raw=C:\Users\jiahao.luo\Desktop\mj\新建文件夹 (2),<这里是数据集路径，把筛选期，后续的访视评估，总评数据集都放这里>
	out=C:\Users\jiahao.luo\Desktop\mj\新建文件夹 (3),<这里是Excel文件输出路径>
    DATEV0=_1recist1,<筛选期肿瘤评估数据集名>
    DATEV=_3recist1,<后续其他肿瘤评估数据集名，如和筛选期在一起留空>
	DATAZP=overall,<肿瘤评估总评数据集名，如没有留空>
	zp=TRGRESP,<肿瘤评估总评字段名，如没有留空>
    V0RS=TRCDAT1,<筛选期肿瘤检查日期>
    VNRS=TRCDAT1,<后续访视肿瘤检查日期>
    V0longest=LDIAM,<筛选期肿瘤最长径>
    VNlongest=LDIAM,<后续访视肿瘤最长径>
    V0sum=SLDIAM,<筛选期肿瘤长径和>
    VNsum=SLDIAM,<后续访视肿瘤长径和>
    V0Location=TULOC1,<筛选期肿瘤部位>
	V0Locations=S_R1SIT,<筛选期肿瘤部位（其他），如没有留空>
    VNLocation=TULOC1,<后续访视肿瘤部位>
	VNLocations=S_R1SIT,<后续访视肿瘤部位（其他），如没有留空>
    V0Record=RecordPosition,<筛选期肿瘤编号，行号>
    VNRecord=RecordPosition<后续访视肿瘤编号，行号>
    Screen=Screening<筛选期访视名称>
	V0TRMETHD=TRMETHD,<筛选期Method of Test or Examination>
	VNTRMETHD=TRMETHD,<后续期Method of Test or Examination>
	V0TRMETHD0=TRMETHD0,<筛选期Method of Test or Examination Specify>
	VNTRMETHD0=TRMETHD0,<后续期Method of Test or Examination Specify>
    );
*改程序默认受试者号编号为Subject
*访视编号为InstanceName
*输出路径为数据集所在路径，格式为EXCEL
*Change history:                                                                                                                        
程序内部QC
√  ×
√已完成程序报错核对           
（暂未完全一致，待确认后调整）已完成程序输出列一致性核对    
√空白输出情况已完成与需求方沟通
√UK和特殊字符情况已处理
********************************************************/

%macro rs(raw=,out=,DATEV0=,DATEV=,DATAZP=,ZP=,V0RS=,VNRS=,V0longest=,V0sum=,V0Location=,V0LocationS=,VNlongest=,VNsum=,VNLocation=,VNLocationS=,V0Record=,VNRecord=,V0TRMETHD=,VNTRMETHD=,V0TRMETHD0=,VNTRMETHD0=,Screen=);
	libname raw "&raw";


data ra;
	set raw.&DATEV0;
	rename &V0RS=TRCDAT1 &V0longest=LDIAM &V0sum=SLDIAM &V0Location=TULOC1 &V0Locations=S_R1SIT &V0Record=RecordPosition;
	keep  Subject InstanceName &V0RS &V0sum &V0longest &V0Location &V0Locations &V0Record;
run;
data rb;
	set raw.&DATEV;
	rename &VNRS=TRCDAT1 &VNlongest=LDIAM &VNsum=SLDIAM &VNLocation=TULOC1 &VNLocations=S_R1SIT &VNRecord=RecordPosition;
	keep  Subject InstanceName &VNRS &VNsum &VNlongest &VNLocation &VNLocations &VNRecord
run;
data r1;
	set ra rb;
run;


data r3;
	set r1;z1=cats(TULOC1,S_R1SIT);
	a=input(compress(TRCDAT1),date9.);
	if a='' then a=input(compress(TRCDAT1),date8.);
run;
proc sort data=r3 ;by Subject InstanceName;run;
proc sort data=r3;by Subject a RecordPosition;run;
data r2;
	set r3;
	retain minSLDIAM; 
	if first.Subject then minSLDIAM=compress(SLDIAM); 
	else minSLDIAM=compress(min(SLDIAM,minSLDIAM)); 
	by Subject ;
	cz=SLDIAM/minSLDIAM;
run;
data r4;
	set r3 ;z2=1;
	if InstanceName='&Screen';
	keep Subject z1 z2;

run;
data r3;
	set r2;
	if InstanceName="&Screen";
	keep Subject SLDIAM;
	rename SLDIAM=SLDIAMa;
run;

proc sort data=r4;by Subject z1;run;
proc sort data=r3;by Subject;run;
proc sort data=r2;by Subject z1;run;
data r2;
	merge r2  r4;
	by Subject z1;
run;
data r2;
	merge r2  r3;
	by Subject;run;
	data f1;set r2;
	cz1=SLDIAM/SLDIAMa;
	if cz>1.2 and abs(SLDIAM-minSLDIAM)>5 then result='PD';
	if cz<1.2 and cz1>0.7 and SLDIAM^='' and SLDIAMa^='' and InstanceName^="&Screen" then result2='SD';
	if abs(SLDIAM-minSLDIAM)<5 and cz1>0.7 and SLDIAM^='' and SLDIAMa^='' and InstanceName^="&Screen" then result2='SD';

	if cz1<=0.7 and SLDIAM^='' and SLDIAMa^=''  then result1='PR';

	if R1NLM^='' then result=R1NLM;
	q4=SLDIAM-minSLDIAM;
	label q4='The difference between Sum of Diameters for All Target Lesions and Smallest previous Sum of Diameter' result2='Caculated  target response(SD)' result1='Caculated  target response(PR)' result='Caculated  target response(PD)' cz='Caculated Result(based on smallest)' cz1='Caculated Result(based on baseline)' minSLDIAM='Smallest previous Sum of Diameter';
run;
data fcr1;
	set f1;
	keep Subject  TRCDAT1 TULOC1 InstanceName;
	rename TRCDAT1=TRCDAT1cr TULOC1=TULOC1cr InstanceName=InstanceNamecr;
run;
proc sort data=f1 ;by Subject TRCDAT1;run;
proc sort data=fcr1 ;by Subject TRCDAT1cr;run;
proc sql;
create table fcr2 as select  * from fcr1 inner join f1
on fcr1.Subject=f1.Subject;
quit;
data fcr3;
	set fcr2;
	TRCDAT1crs=input(compress(TRCDAT1cr),date8.);
	TRCDAT1s=input(compress(TRCDAT1),date8.);
	if TRCDAT1crs='' then TRCDAT1crs=input(compress(TRCDAT1cr),date9.);
	if TRCDAT1s='' then TRCDAT1s=input(compress(TRCDAT1),date9.);
run;
data fcr3;
	set fcr3;
	if TRCDAT1crs<TRCDAT1s and TULOC1cr=TULOC1 then YN='Yes'  ;
run;
data fcr4;
	set fcr3;
	keep Subject InstanceNamecr TULOC1cr InstanceName TRCDAT1cr;
	if YN='Yes'  ;
run;
data fcr5;
	set fcr4;l=1;
	rename TULOC1cr=TULOC1;
	keep Subject InstanceName TULOC1cr l;
run;
proc sort data=fcr5;by Subject InstanceName TULOC1;run;
proc sort data=f1 nodupkey;by Subject InstanceName TRCDAT1;run;

data f2;
	set f1;
	drop SLDIAMa LDIAM RecordPosition a R1NLM z1 z2 TULOC1 S_R1SIT;
	if TRCDAT1^='';
run;


data zp;
	set raw.&DATAZP;rename &zp=TRGRESP;
	keep Subject InstanceName &zp;
run;
proc sort data=f3 ;by Subject InstanceName ;run;
proc sort data=zp ;by Subject InstanceName ;run;
data f3;format q $200.;
	merge f2 zp(in=x);
	if not x then TRGRESP="-";
	by Subject InstanceName;
	if TRGRESP='PD' and result^='PD' then  q='Inconsistent';
	if TRGRESP^='PD' and result='PD' then  q='Inconsistent';
	if TRGRESP='PR' and result1^='PR' and result^='PD' then  q='Inconsistent';
	if TRGRESP^='PR' and result1='PR' and result^='PD' then  q='Inconsistent';
	if TRGRESP='SD' and result2^='SD' then  q='Inconsistent';
	if TRGRESP^='SD' and result2='SD' then  q='Inconsistent';
	if result1^='' and result^='' then q='Multiple standards';
	if result1^='' and result2^='' then q='Multiple standards';
	if result2^='' and result^='' then q='Multiple standards';
	if TRCDAT1='' and SLDIAM='' then delete;
	label q='Result' TRGRESP='EDC Target Response';
run;
proc sort data=f1 ;by Subject InstanceName TULOC1;run;
data fcr6;
	merge f1 fcr5;
	by Subject InstanceName TULOC1;label l='not disappear'; 
run;
data fcr7;
	set fcr6;format l1 $200.;
	if TULOC1^='' and InstanceName^="&Screen";
	if l='' then l1='YES';keep Subject InstanceName l1 ;if l1^='';
	label l1='Is there any disappearance of lesions during this visit';
run;

proc sort data=r1 out=fcr8 nodupkey;by Subject InstanceName TULOC1;run;
data fcr8;
	set fcr8;
	fa=cats(Subject,InstanceName );
	keep fa Subject InstanceName TULOC1 ;
run;
proc sort data=fcr8;by fa;run;
data fcr9;set fcr8;
	TULOC2=lag(TULOC1);
	if first.fa then TULOC2='' ;
	by fa;
run;
data fcr91 fcr92 fcr93;
	set fcr9;format fb1 fb2 fb3 $200.;
	if index(TULOC2,'Lymph Node') and  TULOC2^=''  then output fcr91;
	if index(TULOC1,'Lymph Node') and  TULOC2^=''  then output fcr91;

	/*fb='既有淋巴结，又有非淋巴结';*/
	if index(TULOC1,'Lymph Node') and  TULOC2=''  then output fcr92;
	/*fb='只有淋巴结';*/
	if index(TULOC1,'Lymph Node')=0 and  TULOC2=''  then output fcr93;
	/*fb='只有非淋巴结';*/
run;
data fcr91;
	set fcr91;
	fb1='Both non lymph node and lymph node';keep Subject InstanceName fb1;
run;
data fcr92;
	set fcr92;
	fb2='Only lymph node';;keep Subject InstanceName fb2;
run;
data fcr93;
	set fcr93;
	fb3='Only non lymph node';keep Subject InstanceName fb3;
run;

proc sort data=fcr91 nodupkey;by Subject InstanceName;run;
proc sort data=fcr92 nodupkey;by Subject InstanceName;run;
proc sort data=fcr93 nodupkey;by Subject InstanceName;run;

data fcr10;
	merge fcr91 fcr92 fcr93;
	by Subject InstanceName;
	fb=fb1;
	if fb='' then fb=fb2;
	if fb='' then fb=fb3;
	drop fb1 fb2 fb3;
	label fb='Type of lesion in this visit';
run;
proc sort data=fcr10;by Subject InstanceName;run;
proc sort data=f3;by Subject InstanceName;run;
data f3a;merge f3 fcr10 fcr7;
	by Subject InstanceName;
	if fb='Only lymph node' and SLDIAM<10 then result3='CR';
	if fb='Only non lymph node' and l1^='' then result3='CR';
	if fb='Both non lymph node and lymph node' and l1^='' and SLDIAM<10  then result3='CR';
	if  result1='PR' and result3='CR' then  q='';
	if  result1='PR' and result3='CR' then  result1='';
	if TRGRESP='CR' and result3^='CR' then  q='Inconsistent';
	if TRGRESP^='CR' and result3='CR' then  q='Inconsistent';
	label result3='Caculated  target response(CR)';
	if InstanceName='Screening' then TRGRESP='';
	if TRGRESP='-' then q='EDC Target Response is empty';
    if result='' and result1='' and result2='' and result3='' and TRGRESP='NA' then q='';
	if TRCDAT1^='';
run;
proc sql;
create table f3b as select Subject,InstanceName,TRCDAT1,SLDIAM,minSLDIAM,q4,cz,cz1,fb,l1,result,result2,result1,result3,TRGRESP,q from f3a;
quit;
proc export data=f3b
   outfile="&raw\Recist.xlsx"
   dbms=xlsx replace label;
   sheet="Recist";
run;

libname raw "&raw";

data ra;set raw.&DATEV0;
rename &V0Location=TULOC1  V0Locations=S_R1SIT &V0Record=RecordPosition &V0TRMETHD=TRMETHD1 &V0TRMETHD0=TRMETHDO;
run;
data rb;set raw.&DATEV;
rename &VNLocation=TULOC1  VNLocations=S_R1SIT &VNRecord=RecordPosition &VNTRMETHD=TRMETHD1 &VNTRMETHD0=TRMETHDO;

run;

data ra;
set ra;
TULOC9=compress(cat(TULOC1,compress(S_R1SIT,,'ak')),'	');
TRMETHDF=compress(cat(TRMETHD1,compress(TRMETHDO,,'ak')),'	');
if TRMETHDO^='' then TRMETHDF=compress(cat(TRMETHD1,'(',compress(TRMETHDO),')'));
keep RecordPosition Subject InstanceName TULOC1 S_R1SIT TRMETHD1 S_R1SIT TULOC9 TRMETHDF;
run;
data rb;
set rb;
TULOC9=compress(cat(TULOC1,compress(S_R1SIT,,'ak')),'	');
TRMETHDF=compress(cat(TRMETHD1,compress(TRMETHDO,,'ak')),'	');

if TRMETHDO^='' then TRMETHDF=compress(cat(TRMETHD1,'(',compress(TRMETHDO),')'));
rename TRMETHD1=TRMETHD3;
keep RecordPosition Subject InstanceName TULOC1 S_R1SIT TRMETHD1 S_R1SIT TULOC9  TRMETHDF;
/*rename TRMETHD3=TRMETHD1;*/
run;
data r1;
set ra rb;
y2=cats(Subject,TULOC9,S_R1SIT);
y1=cats(Subject,InstanceName);
xq=UPCASE(compress(TULOC1,,'nk'));


xq1=UPCASE(compress(S_R1SIT,,'nk'));

run;
proc sort data=r1;by y1 xq xq1  RecordPosition;run;
data rq1;set r1;
retain w;
if first.y1 then w=1;
else w=w+1;

by y1;
if RecordPosition=0 then w=0;run; 
data r1;set rq1;
RecordPositiona=RecordPosition;
label RecordPositiona='yyy';
RecordPosition=w;run;
proc sort data=r1;by y2;run;
data rr1;set r1;y1=cats(Subject,InstanceName);
keep Subject InstanceName RecordPosition y1 RecordPositiona;
run;
proc sort data=rr1;by y1 RecordPosition;
run;
data  rr2;
set rr1;
if last.y1;
by y1;rename RecordPosition=RecordPosition1;
keep Subject InstanceName RecordPosition RecordPositiona;

label RecordPosition='Target Number';
run;
data rrq2;set rr2;
if InstanceName="&Screen";rename RecordPosition1=RecordPosition11;
keep  Subject RecordPosition1 RecordPositiona;
run;

proc sort data=rr2;
by Subject;
run;
proc sort data=rrq2;by Subject;run;
data rr3;merge  rr2 rrq2;
by Subject;
if RecordPosition11^=RecordPosition1 and RecordPosition11^='' then q1='Amount Inconsistent';
run;
data rr3;set rr3;
keep q1 RecordPosition11 Subject InstanceName;
run;
proc sort data=rr3;
by Subject InstanceName;
run;
proc sort data=r1;
by Subject InstanceName;
run;
data r2;merge r1 rr3;
by Subject InstanceName;
run;
proc sort data=r2;
by y1;
run;
data r2;
set r2;
y2=lag(TRMETHDF);
if first.y1 then y2='';
by y1;run;
data ff1;set r2;
if y2^=TRMETHDF and y2^='' then q='Method Inconsistent';

label q='Method Result' y2='The Last Method' q1='Number Result' ;
drop y1 y2;
run;
data ff2;set ff1;
if InstanceName="&Screen";
rename TRMETHDF=TRMETHD11;
label TRMETHDF='Screening Method';
keep  Subject TULOC1 S_R1SIT TRMETHDF TULOC9 RecordPosition RecordPosition11 RecordPositiona ;
run;
data ff3;set ff1;
if InstanceName^="&Screen";TULOC11=TULOC1;S_R1SIT1=S_R1SIT ;
drop q TULOC1 S_R1SIT ;
run;
data ff31;set ff3;
if TULOC9='' ;keep Subject InstanceName RecordPosition RecordPosition11 q1 RecordPositiona;
run;
data ff3;set ff3;if  TULOC9='' then delete;
run;
proc sort data=ff2;by  Subject RecordPosition;run;
proc sort data=ff3;by  Subject RecordPosition;run;
proc sort data=ff31;by  Subject ;run;
data ff21;merge   ff2 ff31;by Subject;run;
data ff4;merge   ff21 ff3;format q $200.;
by  Subject RecordPosition;

run;

proc sort data=ff4;by Subject;run;
data new1;set ff4;
if TULOC92^=TULOC91 and TULOC92^='' and TULOC91^='';
run;
data new2;
set ff4;
if TULOC92=TULOC91 then c=1 ;
if TULOC92=''  then c=1 ;
if TULOC91='' then c=1 ;
if c=1;
run;
data new3;
set new1;
rename TULOC92=TULOC93;
keep Subject InstanceName RecordPosition TRMETHD3 TULOC92 TRMETHDF TRMETHD3 S_R1SIT1 TULOC11;
run;
data new4;
set new1;
rename TULOC91=TULOC93;
keep Subject InstanceName RecordPosition TRMETHD1 TULOC91 TRMETHDF TRMETHD3 TRMETHD11 TULOC1 S_R1SIT;
run;
proc sort data=new3;by Subject TULOC93 RecordPosition;run;
proc sort data=new4;by Subject TULOC93 RecordPosition;run;
data new5;
merge new3 new4;
by Subject TULOC93 ;
run;
data ff4;set new2 new5;
/*if InstanceName=''  and sum(Subject)=1 then h=1;*/
if TULOC11='' and TULOC1='' and InstanceName='' then delete;
if TRMETHD11^=TRMETHDF then q='Method Inconsistent';

if TRMETHD11='' then q='Add new';
if TRMETHDF='' then q='Treatment period Method is empty';
if TULOC11='' and TULOC1='' then delete ;
label TRMETHDF='Method of Test or Examination' RecordPosition='Treatment period Record number' RecordPosition11='Screening sum of lesion' q='Method Result' TULOC1='Screening Location' S_R1SIT='Screening Location Within Site Specification' TULOC11='Treatment period Location' S_R1SIT1='Treatment period Location Within Site Specification';
drop TULOC9;
;run;
data ff41;set ff4;
retain k;
if first.Subject then k=1;
else k=k+1;
by Subject;

if last.Subject;keep Subject k;run;
proc sort data=ff41;by Subject;run;
data ff42;
merge ff4 ff41;
by Subject;
if InstanceName='' and k=1 then h=1;
run;
proc sql;
create table ff5 as select Subject,TULOC1,S_R1SIT,TRMETHD11,RecordPosition11,InstanceName,RecordPosition,TULOC11,S_R1SIT1,TRMETHDF,q1,q from ff4;
quit;
data ff6;set ff5;r1=cats(Subject,InstanceName);run;
proc sort data=ff6;by r1;run;
data ff6;set ff6;
	InstanceName1=lag(InstanceName);
	TULOC111=lag(TULOC11);
	if first.r1 then InstanceName1='';
	if first.r1 then TULOC111='';
	by r1;
run;
data ff7;set ff6;
	if InstanceName1=InstanceName and TULOC111=TULOC11 and TULOC11='' then InstanceName='';
	drop InstanceName1 TULOC111 r1;
run;
data rb1;
	set rb;rename TULOC1=wc1;
	keep Subject TULOC1 InstanceName;
if TULOC1^='';
run;
data ff71;set ff7;
wc2=cats(Subject,InstanceName);
if InstanceName^='';
run;
proc sort data=ff71 ;by wc2 RecordPosition;run;

data rf;set ff71;
if last.wc2;
by wc2;
rename RecordPosition=RecordPositionwc;
label RecordPosition='Sum of Treatment period Record number';
keep Subject InstanceName  RecordPosition;
run;
proc sort data=rf ;by Subject InstanceName;run;

proc sort data=rb1 nodupkey;by Subject InstanceName;run;
proc sort data=ff7 ;by Subject InstanceName;run;
data ff8;
	merge ff7 rb1 rf;
	by Subject InstanceName;
	if q='Treatment period Method is empty' and wc1='' then q='No Follow up visits for the subjects';
	if q='Method Inconsistent' and RecordPositionwc^=RecordPosition11 then q='Sum of Record number Inconsistent';
	if q1^='' then q=q1;	
	drop wc1;
run;
proc sql;
create table ff9 as select Subject,TULOC1,S_R1SIT,TRMETHD11,RecordPosition11,InstanceName,RecordPosition,RecordPositionwc,TULOC11,S_R1SIT1,TRMETHDF,q from ff8;
quit;

%let nowdate = %sysfunc(date(), yymmddn8.);
%let xlsxname=Recist&nowdate.;
%put &xlsxname;

proc export data=f3b
   outfile="&out./&xlsxnam..xlsx"
   dbms=xlsx replace label;
   sheet="Recist";

proc export data=ff9
   outfile="&out./&xlsxnam..xlsx"
   dbms=xlsx replace label;
   sheet="Consistency";
run;

%mend;

%RS(raw=C:\Users\jiahao.luo\Desktop\mj\新建文件夹 (2)\5083,
	out=C:\Users\jiahao.luo\Desktop\mj\新建文件夹 (3)\5083,
    DATEV0=_1recist1,
    DATEV=_3recist1,
	DATAZP=recist2,
	zp=R3IORESP,
    V0RS=TRCDAT1,
    VNRS=TRCDAT1,
    V0longest=LDIAM,
    V0sum=SLDIAM,
    V0Location=TULOC1,
	V0Locations=S_R1SIT,
    VNlongest=LDIAM,
    VNsum=SLDIAM,
    VNLocation=TULOC1,
	VNLocations=S_R1SIT,
    V0Record=RecordPosition,
    VNRecord=RecordPosition,
	V0TRMETHD=TRMETHD1,
	VNTRMETHD=TRMETHD3,
	V0TRMETHD0=TRMETHDO,
	VNTRMETHD0=TRMETHDO,
	Screen=Screening
    ); 
