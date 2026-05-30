/**************************************************************
* Program:      dcode
* Function:     编码QC程序;
* Date:         20221230	
* Programmer:   luojiahao/tangyueping
**
**
%dcode(xm=ZGJAK018,path=C:\Users\jiahao.luo\Desktop\新建文件夹 (2)\20221201-删除UID,mc=_c7_,llt=_c8_,lltc=_c9_,pt=_c10_,ptc=_c11_,ymc=_c8_,syz=_c9_,bm=_c17_,mcbm=_c18_,atc4=_c26_);
打开SOP标准编码表格，复制上方语句，修改项目号，path输入生成结果EXCEL的位置（无需提前创建）运行
有问题的数据可筛选每个sheet最后一列
mc,llt,lltc,pt,ptc分别对应报告名称，LLT,LLTCODE,PT,PTCODE位置
ymc,syz,bm,mcbm,atc4分别对应药物报告名称，适应症，药物名称编码，首选名称，ATC4编码
* Change history:
更新了变量对应关系修改
mc,llt,lltc,pt,ptc分别对应报告名称，LLT,LLTCODE,PT,PTCODE位置
ymc,syz,bm,mcbm,atc4分别对应药物报告名称，适应症，药物名称编码，首选名称，ATC4编码

***************************************************************/
/*fin1-ae1;AE 报告名称相同，LLTcode不一致*/
/*fin2-ae2:AE 报告名称相同，PTCode不一致*/
/*fin3-ae3:AE LLT相同，LLTcode不一致*/
/*fin4-ae4:AE LLT=PT，但LLTcode和PTcode不一致*/
/*fin5-cm1:CM 报告名称相同，药物名称编码不一致*/
/*fin6-cm2:CM 报告名称相同，药物首选名称不一致*/
/*fin7-cm3:CM 药物名称编码一致，且用药目的一致，ATC选择不一致*/


%macro dcode(xm=,path=,mc=_c7_,llt=_c8_,lltc=_c9_,pt=_c10_,ptc=_c11_,ymc=_c8_,syz=_c9_,bm=_c17_,mcbm=_c18_,atc4=_c26_);
/*MedDRA*/
%xls2sas(insheet=表格-DM-13-01,outset=AE,row=200000);

/*WHODD*/
%xls2sas(insheet=表格-DM-13-02,outset=cm,row=200000);
data ae1;set ae;
if _c1_^="&xm" then delete;
&mc=compress(&mc,'');
&mc=tranwrd(&mc,'（','(');&mc=tranwrd(&mc,'）',')');/*括号转化*/
run; 
data ae1;set ae1;label _c1_='医学编码报告';run;
/*C8报告名称，C9为LLT，C10为LLT编码，C11为PT，C12为PT编码，模板不同请更改，下同*/
proc sort data=ae1 out=a1;by &mc &lltc;run;
data a2;set a1; 
bg=lag(&mc);
lltcode=lag(&lltc);
run;
data a3;set a2;
if &mc=bg and &lltc^=lltcode then fin='报告名称相同，LLTcode不一致，请核实';
run;
/*1.报告名称相同，LLTcode不一致，请核实*/


proc sort data=ae1 out=b1;by &mc &ptc;run;
data b2;set b1; 
bg=lag(&mc);
ptcode=lag(&ptc);
run;
data b3;set b2;
if &mc=bg and &ptc^=ptcode then fin='报告名称相同，PTCode不一致，请核实';
/*if fin^='';*/
run;
/*2.报告名称相同，PTCode不一致，请核实*/


proc sort data=ae1 out=c1;by &llt &lltc;run;
data c2;set c1; 
llt=lag(&llt);
lltcode=lag(&lltc);
run;
data c3;set c2;
if &llt=llt and &lltc^=lltcode then fin='LLT相同，LLTcode不一致，请核实';
/*if fin^='';*/
run;
/*3.LLT相同，LLTcode不一致，请核实*/


proc sort data=ae1 out=d1;by &llt &pt;run;
data d2;set d1;
if &llt=&pt and &lltc^=&ptc then fin='若LLT=PT，但LLTcode和PTcode不一致，请核实';
/*if fin^='';*/
run;
/*4.若LLT=PT，但LLTcode和PTcode不一致，请核实*/





data cm1;set cm;if _c1_^="&xm" then delete; /*和项目号对应，不同项目请更改*/
&ymc=compress(&ymc,'');/*去除空格*/
&ymc=tranwrd(&ymc,'（','(');&ymc=tranwrd(&ymc,'）',')');/*括号转化*/
run;
data cm1;set cm1;label _c1_='医学编码报告';run;
/*C8报告名称，C9适应症，C17为药物名称编码，C18为首选名称，C19为首选名称编码，模板不同请更改，下同*/
proc sort data=cm1 out=e1;by &ymc &bm;run; 
data e2;set e1; 
bg=lag(&ymc);
drugcode=lag(&bm);
run;
data e3;set e2;
if &ymc=bg and &bm^=drugcode then fin='报告名称相同，药物名称编码不一致，请核实';
/*if fin^='';*/
run;
/*5.报告名称相同，药物名称编码不一致，请核实*/


proc sort data=cm1 out=f1;by &ymc &mcbm;run; 
data f2;set f1; 
bg=lag(&ymc);
pt=lag(&mcbm);
run;
data f3;set f2;
if &ymc=bg and &mcbm^=pt then fin='报告名称相同，药物首选名称不一致，请核实';
/*if fin^='';*/
run;
/*6.报告名称相同，药物首选名称不一致，请核实*/


proc sort data=cm1 out=g1;by &bm &syz &atc4;run; 
data g2;set g1; 
drugcode=lag(&bm);
cmindc=lag(&syz);
atc4=lag(&atc4);
run;
data g3;set g2;
if &bm=drugcode and &syz=cmindc and &atc4^=atc4 then fin='药物名称编码一致，且用药目的一致，ATC选择不一致，请核实';
/*if fin^='';*/
run;
/*7.药物名称编码一致，且用药目的一致，ATC选择不一致，请核实*/



data fin1;set a3;drop bg lltcode;run;
data fin2;set b3;drop bg ptcode;run;
data fin3;set c3;drop llt lltcode;run;
data fin4;set d2;run;
data fin5;set e3;drop bg drugcode;run;
data fin6;set f3;drop bg pt;run;
data fin7;set g3;drop drugcode cmindc atc4;run;
/*%sas2xls(inset=fin1,outsheet=ae1);*/
/*%sas2xls(inset=fin2,outsheet=ae2);*/
/*%sas2xls(inset=fin3,outsheet=ae3);*/
/*%sas2xls(inset=fin4,outsheet=ae4);*/
/*%sas2xls(inset=fin5,outsheet=cm1);*/
/*%sas2xls(inset=fin6,outsheet=cm2);*/
/*%sas2xls(inset=fin7,outsheet=cm3);*/
libname xls EXCEL "&path\a.xls";
data xls."ae1"n(dblabel=YES);set fin1;run;
data xls."ae2"n(dblabel=YES);set fin2;run;
data xls."ae3"n(dblabel=YES);set fin3;run;
data xls."ae4"n(dblabel=YES);set fin4;run;
data xls."cm1"n(dblabel=YES);set fin5;run;
data xls."cm2"n(dblabel=YES);set fin6;run;
data xls."cm3"n(dblabel=YES);set fin7;run;
libname xls clear;
%mend;




