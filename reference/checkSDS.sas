

/**************************************************************
* Program:      checkSDS
* Function:    核查SDS是否合理
* Date:         2024/11/18
* Programmer:   Weijian Zhang
*
* Clinflash Applications for Clinical Data Management
* 
* Sample Usage:

  %checkSDS(
    insds=C:\Users\weijian.zhang\Desktop\SDS.xlsx,    
    inlab=C:\Users\weijian.zhang\Desktop\lab.xlsx,   
    outpath=C:\Users\weijian.zhang\Desktop);          
  


* All Right Reserved 
*
* Change history:
*字典：按照“”DataDictionaryID”，”EntryOID”排序后，“ItemDataString”不为空，且下一行不等于上一行。(去无效化之后)*
*字典：如果DataDictionaryID列包含LBSIG，但是ClinicalType和ClinicalQueryString为空*
*字段：FieldNo大于8位*
*字段：FieldName大于40个字符*
*字段：ControlType=日期框，Unit为空*
*字段：IsGrid=1，HeaderText或Label为空*
*字段：ControlType=单选，多选，下拉时DataDictionaryID为空*
*字段：ControlType=文本标签时ReviewGroups不为空*
*字段：ControlType=文本标签时RequireVerification=1*
*字段：ControlType≠单选，多选，下拉，且IsLab=0时DataDictionaryID不为空*
*字段：ControlType≠文本标签时ReviewGroups为空*
*字段：ControlType≠文本标签时RequireVerification=0*
*字段：DataFormat=yyyy-MM-dd时，IsFutureDateTime不为1*
*字段：FieldNo不等于VariableNo*
*2022-6-9：增加了核查CheckStep和CheckAction无效化数据点*
*2022-6-23：更新了核查FieldName从大于50字符修改为大于40个字符*
*2023-11-30：更新了因sds的字段sheet里面ReviewGroups和ReviewGroups报错问题*
*2024-08-29：更新了因季度更新，变更了SDS*
*2024-11-15：增加了核查实验室指标是否被重复使用*
*2024-11-15：增加了核查实验室指标的名称是否和字段名称一致*
*2025-04-02：增加了核查formoid和SAStext是否一致*
*2025-06-13：增加了核查核查是否填写入组出组*
*2025-06-24：增加了核查核查DM不可见但勾选审核/CRA不可见但勾选校对/PI不可见但勾选签名/核查重复的变量编号*
*2025-10-17：增加了核查核查实验室指标是否勾选结果有值临床评估必填*

**************************************************************/





%macro checkSDS(
    insds=, /* 输入SDS文件的路径 */
	inlab=, /* 输入实验室配置文件的路径 */
    outpath= /* 输出Excel文件的路径 */
);

%let runDate = %sysfunc(date(), yymmdd10.);/* 获取当前日期并格式化为YYYYMMDD格式 */


proc import datafile="&insds" /* 文件路径 */
    out=Dic dbms=xlsx replace; /* 输出数据集名和文件类型 */
    sheet="DatadictionaryEntry"; /* 指定工作表名 */
    getnames=yes; /* 是否从文件首行读入变量名 */
run;

proc import datafile="&insds" /* 文件路径 */
    out=ic dbms=xlsx replace; /* 输出数据集名和文件类型 */
    sheet="DataDictionary"; /* 指定工作表名 */
    getnames=yes; /* 是否从文件首行读入变量名 */
run;


proc import datafile="&insds" /* 文件路径 */
    out=Field dbms=xlsx replace; /* 输出数据集名和文件类型 */
    sheet="Field"; /* 指定工作表名 */
    getnames=yes; /* 是否从文件首行读入变量名 */
run;

proc import datafile="&insds" /* 文件路径 */
    out=VWS dbms=xlsx replace; /* 输出数据集名和文件类型 */
    sheet="VisitWindowSetting"; /* 指定工作表名 */
    getnames=yes; /* 是否从文件首行读入变量名 */
run;



proc import datafile="&inlab" /* 文件路径 */
    out=Lab dbms=xlsx replace; /* 输出数据集名和文件类型 */
    sheet="LabKey"; /* 指定工作表名 */
    getnames=yes; /* 是否从文件首行读入变量名 */
run;

/*数据处理*/
data dic1;set dic;
if isActive='1';
run;
/*去无效化*/
proc sort data=dic1;by dataDictionaryOID entryOID;run;
data dic2;set dic1;
if itemDataString^='';
lagc1=lag(dataDictionaryOID);
lagc2=lag(entryOID);
lagc3=lag(ordinal);
lagc4=lag(itemDataString);
lagc5=lag(isSpecify);
lagc6=lag(isActive);
lagc7=lag(clinicalType);
lagc8=lag(labCommentAvailable);
run;
data fin1;set dic2;
if lagc1=dataDictionaryOID and lagc2=entryOID and lagc3=ordinal and lagc4=itemDataString and lagc5=isSpecify and lagc6=isActive and lagc7=clinicalType and lagc8=labCommentAvailable;
drop lagc1 lagc2 lagc3 lagc4 lagc5 lagc6 lagc7 lagc8;
run;
/*字典：按照“”DataDictionaryID”，”EntryOID”排序后，“ItemDataString”不为空，且下一行不等于上一行。(去无效化之后)*/


data ic2;set ic;
if isClinical='1';
keep dataDictionaryOID isClinical;
rename isClinical=q;
run;
data ic3;
merge dic(in=a) ic2;
if a;
by dataDictionaryOID;
run;
data fin2;
set ic3;
if q='1' and (clinicalType='' or clinicalQueryString='');
drop q;
run;
/*如果DataDictionaryID列包含LBSIG，但是ClinicalType和ClinicalQueryString为空*/


data fin3;set field;
if length(fieldOID)>8;
run;
/*FieldNo大于8位*/


data fin4;set field;
if length(fieldName)>40;
run;
/*FieldName大于40个字符*/


data fin5;set field;
if fieldName='日期框' and unit='';
run;
/*ControlType=日期框，Unit为空*/


data fin6;set field;
where label='';
run;
/*IsGrid=1，Label为空*/


data fin7;set field;
if (controlType='水平单选框' or controlType='垂直单选框' or controlType='多选框') and dataDictionaryOID='';
run;
/*ControlType=单选，多选，下拉时DataDictionaryID为空*/


data fin8;set field;
if controlType='文本标签' and reviewGroups^='';
run;
/*ControlType=文本标签时ReviewGroups不为空*/


data fin9;set field;
if controlType='文本标签' and requireVerification='1';
run;
/*ControlType=文本标签时RequireVerification=1*/



data fin10;set field;
if controlType^='水平单选框' and  controlType^='垂直单选框' and controlType^='多选框' and controlType^='下拉框' and isLab='0' and dataDictionaryOID^='';
run;
/*ControlType≠单选，多选，下拉，且IsLab=0时DataDictionaryID不为空*/


data fin11;set field;
if controlType^='文本标签' and reviewGroups='';
run;
/*ControlType≠文本标签时ReviewGroups为空*/


data fin12;set field;
if controlType^='文本标签' and requireVerification='0';
run;
/*ControlType≠文本标签时RequireVerification=0*/


data fin13;set field;
if index(dataFormat,'yyyy-MM-dd') and isFutureDate^='1';
run;
/*DataFormat=yyyy-MM-dd时，IsFutureDateTime不为1*/


data fin14;set field;
if controlType^='文本标签' and fieldOID^=variableNo;
run;
/*FieldNo不等于VariableNo*/


/*核查实验室指标编号是否等于字段编号*/
data fin15;set field;
where fieldOID^=labKey and labKey^='';
run;




/*核查实验室指标是否被重复使用*/
proc sort data=field nouniquekey out=fin17;
by labKey;
where labKey^='';
run;

/*核查实验室指标名称是否和字段名称一样*/

data lb1;set field;where labKey^='';keep formOID fieldOID fieldName labKey;run;
data lb2;set lab;keep LabKey KeyDescription;run;
proc sort data=lb1;by labKey;run;
proc sort data=lb2;by labKey;run;
data lb3;merge lb1(in=a) lb2;by labKey;if a;run;
data lb4;set lb3;if fieldName^=KeyDescription then q=1;run;
data fin18;set lb4;where q=1;drop q;run;


/*2025-04-02：增加了核查formoid和SAStext是否一致*/
data fin20;set field;
where formOID^=SASText;
run;

/*核查默认值表格勾选了可添加行*/
data fin21;
set field;
if gridDefaultValueDictionary^='' and isGridCanAddRow='1';run;


/*核查是否填写入组出组*/
data fin22;
set VWS;
where inGroup='' or outGroup='';run;

/*DM不可见但勾选审核*/
data fin23;
set field;
if index(viewRestrictions,'DM') and  index(reviewGroups,'DM Review');run;

/*CRA不可见但勾选校对*/
data fin24;
set field;
if index(viewRestrictions,'CRA') and  requireVerification='1';run;

/*PI不可见但勾选签名*/
data fin25;
set field;
if index(viewRestrictions,'PI') and isRequireSign='1';run;

/*核查重复的变量编号*/
data VNO;
set field;
if  variableNo^='';run;

proc sort data=VNO nouniquekey out=fin26;by variableNo;run;

/*实验室指标是否勾选结果有值临床评估必填*/
data fin27;
set field;
if isLab='1' and isClinicalRequired='0';
run;




libname xls EXCEL "&outpath\核查结果&runDate..xlsx";
data xls."ItemDataString不为空"n(dblabel=YES);set fin1;run;
data xls."ClinicalType和QueryString为空"n(dblabel=YES);set fin2;run;
data xls."FieldNo大于8位"n(dblabel=YES);set fin3;run;
data xls."FieldName大于40个字符"n(dblabel=YES);set fin4;run;
data xls."日期框Unit为空"n(dblabel=YES);set fin5;run;
data xls."Label为空"n(dblabel=YES);set fin6;run;
data xls."DataDictionaryID为空"n(dblabel=YES);set fin7;run;
data xls."文本标签时ReviewGroups不为空"n(dblabel=YES);set fin8;run;
data xls."文本标签时勾选需校对"n(dblabel=YES);set fin9;run;
data xls."IsLab等于0时数据字典不为空"n(dblabel=YES);set fin10;run;
data xls."非文本标签时审核为空"n(dblabel=YES);set fin11;run;
data xls."非文本标签时未勾选需校对"n(dblabel=YES);set fin12;run;
data xls."日期格式IsFutureDateTime不为1"n(dblabel=YES);set fin13;run;
data xls."FieldNo不等于VariableNo"n(dblabel=YES);set fin14;run;
data xls."字段编号不等于实验室编号"n(dblabel=YES);set fin15;run;
data xls."实验室指标重复使用"n(dblabel=YES);set fin17;run;
data xls."实验室指标名称不等于字段名称"n(dblabel=YES);set fin18;run;
data xls."formoid和SAStext不一致"n(dblabel=YES);set fin20;run;
data xls."默认值表格勾选了可添加行"n(dblabel=YES);set fin21;run;
data xls."访视窗出入组信息为空"n(dblabel=YES);set fin22;run;
data xls."DM不可见但勾选审核"n(dblabel=YES);set fin23;run;
data xls."CRA不可见但勾选校对"n(dblabel=YES);set fin24;run;
data xls."PI不可见但勾选签名"n(dblabel=YES);set fin25;run;
data xls."核查重复的变量编号"n(dblabel=YES);set fin26;run;
data xls."未勾选结果有值临床评估必填"n(dblabel=YES);set fin27;run;
libname xls clear;
%mend checkSDS;











