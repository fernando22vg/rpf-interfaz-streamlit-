object oLine, objeto, SetSelect, oTerm, oLdf;
set sLine, objetos, sTerm, sTerini;
string sep, numberformat, name;
int fla,col, ierr, num1,num2,num3,num4;
double maxload,iClass;

xlStart();
objetos=GetCalcRelevantObjects('*.ElmLne,*.ElmTr2,*.ElmSym,*.ElmNec.*.ElmGenStat');
SetSelect=GetFromStudyCase('General Set.SetSelect');
sLine=SetSelect.All();
oLdf=GetCaseCommand('ComLdf');
sTerini=AllRelevant('*.ElmTerm');

sep=xlGetDecimalSeparator();
numberformat=sprintf('0%s000', sep);
xlNewWorkbook();
xlAddWorksheet('Sobretensiones');
xlAddWorksheet('Subtensiones');
xlAddWorksheet('Sin Convergencia');
xlAddWorksheet('Sobrecargas');

fla=1;
col=1;
xlActivateWorksheet(1);
xlSetValue(col,fla,'Nombre del componente');
col+=1;
xlSetValue(col,fla,'Nombre de la contingencia');
col+=1;
xlSetValue(col,fla,'Carga en contingencia[%]');
col+=1;
xlSetValue(col,fla,'Maxima sobrecarga permitida[%]');

xlActivateWorksheet(3);
col=1;fla=1;
xlSetValue(col,fla,'Nombre del componente');
col+=1;
xlSetValue(col,fla,'Nombre de la contingencia');
col+=1;
xlSetValue(col,fla,'Tension en contingencia[p.u.]');

xlActivateWorksheet(4);
col=1;fla=1;
xlSetValue(col,fla,'Nombre del componente');
col+=1;
xlSetValue(col,fla,'Nombre de la contingencia');
col+=1;
xlSetValue(col,fla,'Tension en contingencia[p.u.]');

xlActivateWorksheet(2);
col=1;fla=1;
xlSetValue(col,fla,'Contingencia sin convergencia');

num1=2;num2=2;num3=2;num4=2;
oLdf.Execute();
for(oTerm=sTerini.First();oTerm;oTerm=sTerini.Next())
{
 if(oTerm:m:u>0){
  sTerm.Add(oTerm);
 }
}

EchoOff();

for(oLine=sLine.First();oLine;oLine=sLine.Next())
{
 if(oLine:outserv=0){
  oLine.SwitchOff();
  ierr=oLdf.Execute();
  if(ierr<>0){
  xlActivateWorksheet(2);
  fla=num4;col=1;
  xlSetValue(col,fla,oLine:loc_name);
  fla+=1;num4=fla;
  }
  if(ierr=0){
  for(objeto=objetos.First();objeto;objeto=objetos.Next())
  {
   if(objeto:outserv=0){
   if(objeto:c:loading>100){
    xlActivateWorksheet(1);
    fla=num1;col=1;
    xlSetValue(col,fla,objeto:loc_name);
    col+=1;
    xlSetValue(col,fla,oLine:loc_name);
    col+=1;
    xlSetValue(col,fla,objeto:c:loading);
    maxload=100;
    iClass=objeto.IsClass('*ElmLne');
    if(iClass){maxload=objeto:c:maxload;}
    iClass=objeto.IsClass('*ElmTr2');
    if(iClass){maxload=objeto:c:maxload;}
    iClass=objeto.IsClass('ElmTr3');
    if(iClass){maxload=objeto:c:maxload;}
    col+=1;
    xlSetValue(col,fla,maxload);
    col=1;
    fla+=1;
    num1=fla;
   }
  }
  }

  for(oTerm=sTerm.First();oTerm;oTerm=sTerm.Next())
  {
   if({oTerm:iUsage=0}.and.{oTerm:iEarth=0}.and.{oTerm:uknom>=69}){
    if(oTerm:m:u > 1.05){
    xlActivateWorksheet(4);
    fla=num2;col=1;
    xlSetValue(col,fla,oTerm:loc_name);
    col+=1;
    xlSetValue(col,fla,oLine:loc_name);
    col+=1;
    xlSetValue(col,fla,oTerm:m:u);
    fla+=1;col=1;
    num2=fla;
    }
    if(oTerm:m:u < 0.95){
    xlActivateWorksheet(3);
    fla=num3;col=1;
    xlSetValue(col,fla,oTerm:loc_name);
    col+=1;
    xlSetValue(col,fla,oLine:loc_name);
    col+=1;
    xlSetValue(col,fla,oTerm:m:u);
    fla+=1;col=1;
    num3=fla;
    }
   
   }
  }
 } 
oLine.SwitchOn();
 }
}
EchoOn();

input(name,'Nombre del documento');
xlSaveWorkbookAs(name);
xlTerminate();