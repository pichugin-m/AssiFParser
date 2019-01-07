unit u_assifparser_fstandard;

{$mode objfpc}{$H+}

//************************************************************
 //
 //    Модуль u_assifparser_fstandard
 //    Copyright (c) 2016  Pichugin M.
 //    This file is part of the Assi System.
 //
 //    Разработчик: Pichugin M. (e-mail: pichugin_m@mail.ru)
 //
 //    (2019-01-02)  v.5
 //    (2016-07-05)  v.4
 //
//************************************************************

interface

uses
    Variants, Classes, SysUtils, LazUTF8, Math,
    u_assifparser_class;

{
    ======Стандартные формулы======

    Sum(X1,X2,X3,...,Xn)
    Difference(X1,X2,X3,...,Xn)
    Product(X1,X2,X3,...,Xn)
    Quotient(X1,X2,X3,...,Xn)

    ReplaceEqual(Text,OldText1,NewText1,OldText2,NewText2,...,,OldTextN,NewTextN)
    ReplaceAll(Text,OldText,NewText)
    Replace(Text,OldText,NewText)
    Concat(N1,N2,N3,...,Nn)

    Not(BoolResult)
    If(Check,TrueResult,FalseResult)
    IfErr(Check,TrueResult)
    Equal(N1,N2,N3,...,Nn)
    Num(X1)
    Text(X1)
    Now(TextFormat)
}

procedure SetDefaultFormula(AParser:TAssiFormulaParser);

implementation

{ Стандартные формулы }

//ifErr(Formula,IfErrResult)
procedure Formula_ifErr(AParser:TAssiFormulaParser; var AResult:Variant);
var
  sTmp :String;
begin
  if AParser.GetFormulaParamCount=2 then
  begin
       sTmp:=VarToStr(AParser.GetFormulaParams(0));
       if CompareText(FORMULA_RESULT_ERROR, sTmp)=0 then
       begin
          AResult:=AParser.GetFormulaParams(1);
       end
       else begin
          AResult:=AParser.GetFormulaParams(0);
       end;
  end
  else begin
    AResult:=null;
  end;
end;

//Конвертировать в текст
procedure Formula_ToText(AParser:TAssiFormulaParser; var AResult:Variant);
var
  sTmp :String;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      sTmp:=VarToStr(AParser.GetFormulaParams(0));
      AResult:=sTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Конвертировать в число
procedure Formula_ToNumber(AParser:TAssiFormulaParser; var AResult:Variant);
var
  sTmp :String;
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=0;
      sTmp:=VarToStr(AParser.GetFormulaParams(0));
      sTmp:=StringReplace(sTmp, '.', DecimalSeparator, [rfReplaceAll]);
      sTmp:=StringReplace(sTmp, ',', DecimalSeparator, [rfReplaceAll]);
      sTmp:=StringReplace(sTmp, ' ', '', [rfReplaceAll]);
      if TryStrToFloat(sTmp,dTmp,DefaultFormatSettings) then
         AResult:=dTmp
      else
         AResult:=null;
  end
  else begin
    AResult:=null;
  end;
end;

//Сравнение параметров
procedure Formula_Equal(AParser:TAssiFormulaParser; var AResult:Variant);
var
  i,k,j   :Integer;
  vTmp1,
  vTmp2   :Variant;
begin
  k:=AParser.GetFormulaParamCount;
  j:=0;
  if k>0 then
  begin
       for i:=1 to k-1 do
       begin
           vTmp1:=AParser.GetFormulaParams(i-1);
           vTmp2:=AParser.GetFormulaParams(i);
           //(VarType(vTmp1)=VarType(vTmp2))
           if (VarCompareValue(vTmp1,vTmp2)<>vrEqual) then
           begin
              inc(j);
           end;
       end;
       if j=0 then
         AResult:=True
       else
         AResult:=False;
  end
  else begin
    AResult:=null;
  end;
end;

//Не
procedure Formula_Not(AParser:TAssiFormulaParser; var AResult:Variant);
var
  k    :Integer;
  vTmp1:Variant;
begin
  k:=AParser.GetFormulaParamCount;
  if k=1 then
  begin
     vTmp1:=AParser.GetFormulaParams(0);
     if VarType(vTmp1)=varboolean then
     begin
        AResult:=not vTmp1;
     end
     else begin
        AResult:=null;
     end;
  end
  else begin
    AResult:=null;
  end;
end;

//Или
procedure Formula_Or(AParser:TAssiFormulaParser; var AResult:Variant);
var
  i,k  :Integer;
begin
  k:=AParser.GetFormulaParamCount;
  if k>1 then
  begin
       AResult:=False;
       for i:=0 to k-1 do
       begin
           if AParser.GetFormulaParams(i)=True then
           begin
              AResult:=True;
              break;
           end;
       end;
  end
  else begin
    AResult:=False;
  end;
end;

//И
procedure Formula_And(AParser:TAssiFormulaParser; var AResult:Variant);
var
  i,k,j,l   :Integer;
begin
  l:=AParser.GetFormulaParamCount;
  if l>1 then
  begin
       j:=0;
       k:=0;
       AResult:=False;
       for i:=0 to l-1 do
       begin
           inc(j);
           if AParser.GetFormulaParams(i)=True then
           begin
              inc(k);
           end;
       end;
       if j=k then
          AResult:=True;
  end
  else begin
    AResult:=False;
  end;
end;

//Условие
procedure Formula_If(AParser:TAssiFormulaParser; var AResult:Variant);
var
  k    :Integer;
  vTmp1:Variant;
begin
  k:=AParser.GetFormulaParamCount;
  if k=3 then
  begin
     vTmp1:=AParser.GetFormulaParams(0);
     if VarType(vTmp1)=varboolean then
     begin
        if vTmp1=True then
           AResult:=AParser.GetFormulaParams(1)
        else
           AResult:=AParser.GetFormulaParams(2);
     end
     else begin
        AResult:=null;
     end;
  end
  else begin
    AResult:=null;
  end;
end;

//Условие
procedure Formula_True(AParser:TAssiFormulaParser; var AResult:Variant);
begin
  AResult:=True;
end;

//Условие
procedure Formula_False(AParser:TAssiFormulaParser; var AResult:Variant);
begin
  AResult:=False;
end;

{ Стандартная математика }

//Сложение, Сумма
procedure Formula_Sum(AParser:TAssiFormulaParser; var AResult:Variant);
var
  i,k,a   :Integer;
  vTmp1   :Variant;
  dTmp    :Double;
begin
  k:=AParser.GetFormulaParamCount;
  if k>1 then
  begin
       dTmp:=0;
       a:=0;
       for i:=0 to k-1 do
       begin
           vTmp1:=AParser.GetFormulaParams(i);
           if (VarIsFloat(vTmp1)) then
           begin
              dTmp:=dTmp+vTmp1;
           end
           else begin
               inc(a);
               break;
           end;
       end;

       if a=0 then
         AResult:=dTmp
       else
         AResult:=null;
  end
  else begin
    AResult:=null;
  end;
end;

//Вычитание, Разность
procedure Formula_Difference(AParser:TAssiFormulaParser; var AResult:Variant);
var
  i,k,a   :Integer;
  vTmp1,
  vTmp2   :Variant;
  dTmp    :Double;
begin
  k:=AParser.GetFormulaParamCount;
  if k>1 then
  begin
       vTmp2:=AParser.GetFormulaParams(0);
       if (VarIsFloat(vTmp2)) then
       begin
         dTmp:=vTmp2;
         a:=0;
         for i:=1 to k-1 do
         begin
             vTmp1:=AParser.GetFormulaParams(i);
             if (VarIsFloat(vTmp1)) then
             begin
                dTmp:=dTmp-vTmp1;
             end
             else begin
                 inc(a);
                 break;
             end;
         end;
       end
       else begin
           inc(a);
       end;

       if a=0 then
         AResult:=dTmp
       else
         AResult:=null;
  end
  else begin
    AResult:=null;
  end;
end;

//Умножение, Произведение
procedure Formula_Product(AParser:TAssiFormulaParser; var AResult:Variant);
var
  i,k,a   :Integer;
  vTmp1,
  vTmp2   :Variant;
  dTmp    :Double;
begin
  k:=AParser.GetFormulaParamCount;
  if k>1 then
  begin
       vTmp2:=AParser.GetFormulaParams(0);
       if (VarIsFloat(vTmp2)) then
       begin
         dTmp:=vTmp2;
         a:=0;
         for i:=1 to k-1 do
         begin
             vTmp1:=AParser.GetFormulaParams(i);
             if (VarIsFloat(vTmp1)) then
             begin
                dTmp:=dTmp*vTmp1;
             end
             else begin
                 inc(a);
                 break;
             end;
         end;
       end
       else begin
           inc(a);
       end;

       if a=0 then
         AResult:=dTmp
       else
         AResult:=null;
  end
  else begin
    AResult:=null;
  end;
end;

//Деление, Частное
procedure Formula_Quotient(AParser:TAssiFormulaParser; var AResult:Variant);
var
  i,k,a   :Integer;
  vTmp1,
  vTmp2   :Variant;
  dTmp    :Double;
begin
  k:=AParser.GetFormulaParamCount;
  if k>1 then
  begin
       vTmp2:=AParser.GetFormulaParams(0);
       if (VarIsFloat(vTmp2)) then
       begin
         dTmp:=vTmp2;
         a:=0;
         for i:=1 to k-1 do
         begin
             vTmp1:=AParser.GetFormulaParams(i);
             if (VarIsFloat(vTmp1)) then
             begin
                if vTmp1<>0 then
                begin
                  dTmp:=dTmp/vTmp1;
                end
                else begin
                  AResult:=null;
                  a:=0;
                  break;
                end;
             end
             else begin
                 inc(a);
                 break;
             end;
         end;
       end
       else begin
           inc(a);
       end;

       if a=0 then
         AResult:=dTmp
       else
         AResult:=null;
  end
  else begin
    AResult:=null;
  end;
end;

//Синус
procedure Formula_Sin(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=AParser.GetFormulaParams(0);
      dTmp:=math.degtorad(dTmp);
      dTmp:=sin(dTmp);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//АркСинус
procedure Formula_ArcSin(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=AParser.GetFormulaParams(0);
      dTmp:=arcsin(dTmp);
      dTmp:=math.radtodeg(dTmp);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Косинус
procedure Formula_Cos(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=AParser.GetFormulaParams(0);
      dTmp:=math.degtorad(dTmp);
      dTmp:=cos(dTmp);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//АркКосинус
procedure Formula_ArcCos(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=AParser.GetFormulaParams(0);
      dTmp:=arccos(dTmp);
      dTmp:=math.radtodeg(dTmp);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//АркТангенс
procedure Formula_ArcTan(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=AParser.GetFormulaParams(0);
      dTmp:=arctan(dTmp);
      dTmp:=math.radtodeg(dTmp);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Тангенс
procedure Formula_Tan(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=AParser.GetFormulaParams(0);
      dTmp:=math.degtorad(dTmp);
      dTmp:=tan(dTmp);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Котангенс
procedure Formula_Cot(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=AParser.GetFormulaParams(0);
      dTmp:=math.degtorad(dTmp);
      dTmp:=cot(dTmp);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Корень
procedure Formula_Sqrt(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=AParser.GetFormulaParams(0);
      dTmp:=sqrt(dTmp);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Pi
procedure Formula_Pi(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  dTmp:=Pi;
  AResult:=dTmp;
end;

//Логарифм
procedure Formula_Log(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=AParser.GetFormulaParams(0);
      dTmp:=ln(dTmp);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Логарифм10
procedure Formula_Log10(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=AParser.GetFormulaParams(0);
      dTmp:=log10(dTmp);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Логарифм2
procedure Formula_Log2(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmp :Double;
begin
  if AParser.GetFormulaParamCount=1 then
  begin
      dTmp:=AParser.GetFormulaParams(0);
      dTmp:=log2(dTmp);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Степень(База,Степень)
procedure Formula_Pow(AParser:TAssiFormulaParser; var AResult:Variant);
var
  dTmpB,dTmpE,dTmp :Double;
begin
  if AParser.GetFormulaParamCount=2 then
  begin
      dTmpB:=AParser.GetFormulaParams(0);
      dTmpE:=AParser.GetFormulaParams(0);
      dTmp:=power(dTmpB,dTmpE);
      AResult:=dTmp;
  end
  else begin
    AResult:=null;
  end;
end;

{ Стандартное форматирование }

//Объединение
procedure Formula_Concat(AParser:TAssiFormulaParser; var AResult:Variant);
var
  i,k     :Integer;
  vTmp1   :Variant;
  sTmp    :String;
begin
  k:=AParser.GetFormulaParamCount;
  if k>1 then
  begin
       sTmp:='';
       for i:=0 to k-1 do
       begin
           vTmp1:=AParser.GetFormulaParams(i);
           sTmp:=sTmp+VarToStr(vTmp1);
       end;

       AResult:=sTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Замена содержимого строки при совпадении с одним из вариантов
procedure Formula_ReplaceEqual(AParser:TAssiFormulaParser; var AResult:Variant);
var
  i,k   :Integer;
  sTmp  :String;
begin
  k:=AParser.GetFormulaParamCount;
  if (k>=3)and(VarIsStr(AParser.GetFormulaParams(0))) then
  begin
       i:=1;
       sTmp:=AParser.GetFormulaParams(0); //Текст
       while i<k do
       begin
           if (CompareText(sTmp, AParser.GetFormulaParams(i))=0)
               and(k>i+1) then
           begin
              sTmp:=AParser.GetFormulaParams(i+1);
              break;
           end
           else begin
             i:=i+2;
           end;
       end;
       AResult:=sTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Замена
procedure Formula_Replace(AParser:TAssiFormulaParser; var AResult:Variant);
var
   k   :Integer;
   sTmp:String;
begin
  k:=AParser.GetFormulaParamCount;
  if (k>=3)and(VarIsStr(AParser.GetFormulaParams(0))) then
  begin
       sTmp:=AParser.GetFormulaParams(0);
       sTmp:=StringReplace(sTmp,AParser.GetFormulaParams(1),
           AParser.GetFormulaParams(2),[rfIgnoreCase]);
       AResult:=sTmp;
  end
  else begin
    AResult:=null;
  end;
end;

//Замена
procedure Formula_ReplaceAll(AParser:TAssiFormulaParser; var AResult:Variant);
var
   k   :Integer;
   sTmp:String;
begin
  k:=AParser.GetFormulaParamCount;
  if (k>=3)and(VarIsStr(AParser.GetFormulaParams(0))) then
  begin
       sTmp:=AParser.GetFormulaParams(0);
       sTmp:=StringReplace(sTmp,AParser.GetFormulaParams(1),
           AParser.GetFormulaParams(2),[rfIgnoreCase,rfReplaceAll]);
       AResult:=sTmp;
  end
  else begin
    AResult:=null;
  end;
end;

{Date Time}

//Текущая дата и время
procedure Formula_Now(AParser:TAssiFormulaParser; var AResult:Variant);
var
  k     :Integer;
  vTmp1 :Variant;
begin
  k:=AParser.GetFormulaParamCount;
  if k=0 then
  begin
       AResult:=FormatDateTime(DefaultFormatSettings.ShortDateFormat, now);
  end
  else if k=1 then
  begin
       vTmp1:=AParser.GetFormulaParams(0);
       if VarType(vTmp1)=varstring then
       begin
          AResult:=FormatDateTime(vTmp1,now);
       end
       else begin
          AResult:=null;
       end;
  end
  else begin
    AResult:=null;
  end;
end;

//Test
procedure Formula_Test(AParser:TAssiFormulaParser; var AResult:Variant);
begin
  AResult:='It test message';
end;

{ Прочее }

procedure SetDefaultFormula(AParser:TAssiFormulaParser);
var
   i:integer;
begin
   AParser.FormulaInitNew(33);
   i:=0;
   AParser.FormulaInsert(i,'Text',          @Formula_ToText);
   inc(i);
   AParser.FormulaInsert(i,'Num',           @Formula_ToNumber);
   inc(i);
   AParser.FormulaInsert(i,'Equal',         @Formula_Equal);
   inc(i);
   AParser.FormulaInsert(i,'IfErr',         @Formula_IfErr);
   inc(i);
   AParser.FormulaInsert(i,'If',            @Formula_if);
   inc(i);
   AParser.FormulaInsert(i,'True',          @Formula_True);
   inc(i);
   AParser.FormulaInsert(i,'False',         @Formula_False);
   inc(i);
   AParser.FormulaInsert(i,'Not',           @Formula_Not);
   inc(i);
   AParser.FormulaInsert(i,'Or',            @Formula_Or);
   inc(i);
   AParser.FormulaInsert(i,'And',           @Formula_And);
   inc(i);
   AParser.FormulaInsert(i,'Concat',        @Formula_Concat);
   inc(i);
   AParser.FormulaInsert(i,'Replace',       @Formula_Replace);
   inc(i);
   AParser.FormulaInsert(i,'ReplaceAll',    @Formula_ReplaceAll);
   inc(i);
   AParser.FormulaInsert(i,'ReplaceEqual',  @Formula_ReplaceEqual);
   inc(i);
   AParser.FormulaInsert(i,'Sum',           @Formula_Sum);
   inc(i);
   AParser.FormulaInsert(i,'Difference',    @Formula_Difference);
   inc(i);
   AParser.FormulaInsert(i,'Product',       @Formula_Product);
   inc(i);
   AParser.FormulaInsert(i,'Quotient',      @Formula_Quotient);
   inc(i);
   AParser.FormulaInsert(i,'Cos',           @Formula_Cos);
   inc(i);
   AParser.FormulaInsert(i,'Sin',           @Formula_Sin);
   inc(i);
   AParser.FormulaInsert(i,'ArcCos',        @Formula_ArcCos);
   inc(i);
   AParser.FormulaInsert(i,'ArcSin',        @Formula_ArcSin);
   inc(i);
   AParser.FormulaInsert(i,'Cot',           @Formula_Cot);
   inc(i);
   AParser.FormulaInsert(i,'ArcTan',        @Formula_ArcTan);
   inc(i);
   AParser.FormulaInsert(i,'Tan',           @Formula_Tan);
   inc(i);
   AParser.FormulaInsert(i,'Pi',            @Formula_Pi);
   inc(i);
   AParser.FormulaInsert(i,'Sqrt',          @Formula_Sqrt);
   inc(i);
   AParser.FormulaInsert(i,'Pow',           @Formula_Pow);
   inc(i);
   AParser.FormulaInsert(i,'Log',           @Formula_Log);
   inc(i);
   AParser.FormulaInsert(i,'Log2',          @Formula_Log2);
   inc(i);
   AParser.FormulaInsert(i,'Log10',         @Formula_Log10);
   inc(i);
   AParser.FormulaInsert(i,'Test',          @Formula_Test);
   inc(i);
   AParser.FormulaInsert(i,'Now',           @Formula_Now);
   inc(i);

   AParser.FormulaSynonymAdd('Sum',         'math_s');
   AParser.FormulaSynonymAdd('Difference',  'math_d');
   AParser.FormulaSynonymAdd('Product',     'math_p');
   AParser.FormulaSynonymAdd('Quotient',    'math_q');

   {
   AParser.FormulaSynonymAdd('Sum',         'Сложить');
   AParser.FormulaSynonymAdd('Difference',  'Вычесть');
   AParser.FormulaSynonymAdd('Product',     'Умножить');
   AParser.FormulaSynonymAdd('Quotient',    'Разделить');

   AParser.FormulaSynonymAdd('Sum',         'Сумма');
   AParser.FormulaSynonymAdd('Difference',  'Разность');
   AParser.FormulaSynonymAdd('Product',     'Произведение');
   AParser.FormulaSynonymAdd('Quotient',    'Частное');

   AParser.FormulaSynonymAdd('If',          'Если');
   AParser.FormulaSynonymAdd('Concat',      'Сцепить');
   AParser.FormulaSynonymAdd('Equal',       'Равно');
   }
end;

end.

