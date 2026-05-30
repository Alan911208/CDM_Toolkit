 /**************************************************************
* Program:      dque
* Function:     检查实验室指标;
* Date:         20230613
* Programmer:   luojiahao
*
* Clinflash Applications for Clinical Data Management
* 
* 使用方法:
%dque(raw=C:\Users\jiahao.luo\Desktop\新建文件夹 (2)\20221201-删除UID,dat=lb2,path=C:\Users\jiahao.luo\Desktop\新建文件夹 (2)\20221201-删除UID);
复制上面的语句，raw=输入原始数据集名称，dat=输入需要核查的数据集，path=输出EXECL位置（无需提前创建EXCEL文档）
* Change history:
202212初版
20230613增加了fin5异常有变为异常无的输出
****************************************************************/
/*rawdata位置，具体的数据集*/

%macro dque(raw=,dat=,path=);
libname raw "&raw";

data lb1;set raw.&dat;
keep usubjid siteid visit form LBTEST LBORRES LBORRESU LBORNRLO LBORNRHI LBSIG;run;
data f1;set lb1;
if LBORRES^='' and LBORRESU='' then a='结果不为空，单位为空' ;
if LBORRES^='' and LBORNRLO='' then b='结果不为空，下限为空';
if LBORRES^='' and LBORNRHI='' then c='结果不为空，上限为空';
if index(lborres,'。') or index(lborres,'，') or index(lborres,'阴') or index(lborres,'阳') or index(lborres,'+') or index(lborres,'-') or index(lborres,'±') or index(lborres,'<') or index(lborres,'>') or index(lborres,'≥') or index(lborres,'≤') then b='结果带有"。","，","-","±","阴阳性","<",">","≥","≤","+"';
run;
data fin1;set f1;if a^='' and b^='' and c^='';d='单位，上下限都为空';drop a b c;run;
data lb2;set lb1;
LBORNRHI1=compress(LBORNRHI,'.','kdi');
LBORNRLO1=compress(LBORNRLO,'.','kdi');
run;
data fin2;set lb2;
if LBORNRLO1>LBORRES*10 and LBORRES^='' and LBORNRLO1^='' and  not index(LBSIG,"异常有") then a='结果*10仍小于正常值范围，且临床评估不为异常有临床意义';

if LBORNRHI1*10<LBORRES and LBORRES^='' and LBORNRHI1^='' then a='结果大于正常值范围10倍';
if LBORNRLO1>LBORRES*10 and LBORRES^='' and LBORNRLO1^='' then a='结果*10仍小于正常值范围';
if LBORNRHI1*10<LBORRES and LBORRES^='' and LBORNRHI1^='' and  not index(LBSIG,"异常有") then a='结果大于正常值范围10倍,且临床评估不为异常有临床意义';
if a^='';
run;
data ww;set fin2;sad=compress(input(lborres,10.),'<>=',);run;
proc sort data=ww;by lbtest sad;run;
data ww;set ww;retain w1;if first.lbtest then w1=1;else w1=w1+1;by lbtest;run;
proc sort data=ww;by lbtest descending sad;run;
data ww;set ww;retain w2;if first.lbtest then w2=1;else w2=w2+1;by lbtest;run;
data ww1;set ww;if w1<=3 or w2<=3;run;
data fin2;set ww1;drop w2 w1 sad;run;
proc sort data=lb1;by LBTEST LBORRESU;run;
data lb3;set lb1;
LBORRESU1=lag(LBORRESU);by LBTEST;if first.LBTEST then call missing(LBORRESU1);run;
data lb31;set lb3;if LBORRESU1^='' and LBORRESU^='' and LBORRESU1^=LBORRESU then a='同指标单位不同';run;
proc sort data=lb1;by LBTEST LBORNRLO;run;
data lb3;set lb1;
LBORNRLO1=lag(LBORNRLO);by LBTEST;if first.LBTEST then call missing(LBORNRLO);run;
data lb32;set lb3;if LBORNRLO^='' and LBORNRLO1^='' and LBORNRLO1^=LBORNRLO then b='同指标下限不同';run;
proc sort data=lb1;by LBTEST LBORNRHI;run;
data lb3;set lb1;
LBORNRHI1=lag(LBORNRHI);by LBTEST;if first.LBTEST then call missing(LBORNRHI1);run;
data lb33;set lb3;if LBORNRHI1^='' and LBORNRHI^='' and LBORNRHI1^=LBORNRHI then c='同指标上限不同';run;
proc sort data=lb31;by usubjid visit LBTEST LBORRES;run;
proc sort data=lb32;by usubjid visit LBTEST LBORRES;run;
proc sort data=lb33;by usubjid visit LBTEST LBORRES;run;
data fin3;merge lb31 lb32 lb33;by usubjid visit LBTEST LBORRES;run;
data fin3;set lb31;if a^='';run;
data lb5;set raw.&dat;keep usubjid visit lbtest lborres LBORNRHI LBORNRLO;run;
data lb2;set lb5;jg=compress(lborres,'.','d');run;
data fin4;set lb2;if jg^=''; check='结果有除了数字以外的内容';run;
data a;set raw.&dat;keep usubjid __STUDYEVENTOID visit PAGE LBTEST LBORRES LBORRESN LBSIG LBORRESU LBORNRLO LBORNRHI;
if if index(lbsig,'异常无') or index(lbsig,'异常有');
run;
data a;set a;LBORNRLO=compress(LBORNRLO,'<>=');LBORNRHI=compress(LBORNRHI,'<>=');run;
proc sort data=a;by usubjid LBTEST __STUDYEVENTOID;run;
data b;set a;
LBORRES1=lag(LBORRES);
if first.lbtest then call missing(LBORRES1);
LBSIG1=lag(LBSIG);
if first.lbtest then call missing(LBSIG1);
LBORNRLO1=lag(LBORNRLO);
if first.lbtest then call missing(LBORNRLO1);
LBORNRHI1=lag(LBORNRHI);
if first.lbtest then call missing(LBORNRHI1);
by usubjid lbtest ;
if LBORRES='' then delete;
run;
data fin5;set b;
if index(lbsig,'异常无') and index(lbsig1,'异常有') and LBORRES<LBORRES1<LBORNRLO then query='本次访视结果小于上次访视结果小于正常值下限,临床评估从异常有意义变为异常无意义';
if index(lbsig,'异常无') and index(lbsig1,'异常有') and LBORRES>LBORRES1>LBORNRHI then query='本次访视结果大于上次访视结果大于正常值上限,临床评估从异常有意义变为异常无意义';
if query^='';
run;

libname xls EXCEL "&path\a.xls";
data xls."结果单位上下限都为空"n(dblabel=YES);set fin1;run;
data xls."极值核查"n(dblabel=YES);set fin2;run;
data xls."同实验室指标单位范围核查"n(dblabel=YES);set fin3;run;
data xls."结果有除数字以外的内容"n(dblabel=YES);set fin4;run;
data xls."异常有变为异常无"n(dblabel=YES);set fin5;run;
libname xls clear;
%mend;




