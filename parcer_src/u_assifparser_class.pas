unit u_assifparser_class;

{$mode objfpc}{$H+}

//************************************************************
 //
 //    Модуль u_assifparser_class
 //    Copyright (c) 2019  Pichugin M.
 //    This file is part of the Assi System.
 //
 //    Разработчик: Pichugin M. (e-mail: pichugin_m@mail.ru)
 //
 //    (2019-01-07)
 //
//************************************************************

interface

uses
      Classes, SysUtils, LazUTF8, ComCtrls, Variants, Dialogs;

type

 { Forward Declarartions }

  TAssiFormulaParser = class;
  TOnFormulaError = procedure(Sender:TObject; AMessage:String);

  { Data types }

  TDoFormulaResult = procedure(AParser:TAssiFormulaParser; var AResult:Variant);

  TAFPStyle = (afpsUnknow, afpsText, afpsNum, afpsBool,
              afpsKeyMath, afpsKeyFun, afpsGroupParam, afpsGroup);

  PAFPItem =^TAFPItem;
  TAFPItem = record
    ID            :integer;
    ParentID      :integer;
    StyleInit     :TAFPStyle;//Значение Style при инициализации
    Style         :TAFPStyle;//группа,функция,число
    Name          :ShortString;
    Content       :Variant;
    Result        :Variant;
    ContentAsText :String;
    SubItems      :Array of PAFPItem;
    Ignor         :boolean;
  end;

  { TAssiFormulaParser }

  TAssiFormulaParser = class(TObject)
  private
    FIDCounter       :integer;
    FOnFormulaError  :TOnFormulaError;
    FParsesIndex     :integer;
    FDataCol         :Pointer;
    FDataRow         :Pointer;
    FDataTable       :Pointer;
    FResult          :Variant;
    FArrayFormulaParts           :Array of Array of PAFPItem;
    FArrayFormulaName            :Array of string;
    FArrayFormulaPointer         :Array of Pointer;
    FArraySynonymNameOriginal    :Array of string;
    FArraySynonymNameAlternative :Array of string;
    FArrayFormulaParams          :Array of Variant;
    function DoCheckSequence(AParentItem: PAFPItem): Boolean;
    function GroupStyle(AItem: PAFPItem): integer;
    function GetFormulaIndex(AName:String):Integer;
    //Добавить новый эл-т парсинга
    function AddFPart:PAFPItem;
    //Очистить результаты парсинга
    procedure ClearFParts;
    //Получить тип элемента
    function GetParentStyle(AID:integer):TAFPStyle;
    //Создание стандартных формул
    procedure FormulaAddDefault();
    function GetID:integer;
  protected
    function DoReadString(AFormula: String;
             ALevel:integer=0; AParent:PAFPItem=nil):Boolean;
    function DoCalcFormula(AParentItem :PAFPItem):Boolean;
  public
    property DataTable:Pointer read FDataTable write FDataTable;
    property DataRow:Pointer read FDataRow write FDataRow;
    property DataCol:Pointer read FDataCol write FDataCol;
    property OnFormulaError:TOnFormulaError read FOnFormulaError write FOnFormulaError;
    //Создание нового парсинга с сохранением результатов предыдущего
    function NewParse:Integer;
    //Кол-во парсингов
    function GetParseCount:Integer;
    //Изменение индекса текущего парсинга
    function SetCurrentParse(AIndex:Integer):boolean;
    //Очистка всех парсингов
    function ClearParses:boolean;

    //Проверка зарезервированных имен формул
    function IsReservedName(AName:String):Boolean;
    //Только распарсить формулу
    function Parse(AFormula:String):Boolean;
    //Только выполнить формулу
    function Perform:Boolean;
    //Распарсить и выполнить формулу
    function Execute(AFormula:String):Boolean;
    procedure UploadToTree(ATree: TTreeView;
             ParentNode:TTreeNode=nil; AIDParent:integer=0);
    function GetFormulaParams(AIndex:integer):Variant;
    function GetFormulaParamCount:integer;
    function GetResult:Variant;
    function GetResultAsText:String;
    function GetFormulaCount:Integer;
    procedure FormulaSynonymAdd(AOriginalName, ASynonymName:String);
    //Добавление пользовательской формулы
    procedure FormulaAdd(AName:String; AFunc:Pointer);
    procedure FormulaInitNew(ANewCount:Integer);
    procedure FormulaInsert(Index:integer; AName:String; AFunc:Pointer);
    constructor Create;
    destructor Destroy; override;
  end;

const
      MATHCHARS                    = '/*+-';
      ABCCHARS                     = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
      QUOTE1                       = '&quot;';
      QUOTE2                       = '\"';
      QUOTE3                       = '"';
      BKT_L                        = '(';
      BKT_R                        = ')';
      FORMULA_CHRPARAMSBEGIN       = '(';
      FORMULA_CHRPARAMSEND         = ')';
      FORMULA_CHRPARAMSSEPARETER   = ',';
      FORMULA_CHRPARAMSSEPARETER_Alt2  = '\,';
      FORMULA_CHRPARAMSSEPARETER_Alt1  = '%#044';
      FORMULA_BAFFLE               = '\';
      FORMULA_RESULT_TRUE          = 'True';
      FORMULA_RESULT_FALSE         = 'False';

      FORMULA_RESULT_ERROR         = '#ERROR';
      FORMULA_RESULT_ERRORPARNAME  = '#ERRORPN';


function isCorrectNumber(AData:String):boolean;
function GetCorrectNumber(AData:String):String;

implementation

uses
    u_assifparser_fstandard;

function isCorrectNumber(AData:String):boolean;
const
  CHRDOT1='.';
  CHRDOT2=',';
var
  LStr,
  i,
  CountA,
  CountB,
  CountC    :Integer;
  tmpChrs,
  tmpChr    :String;
begin
  LStr  :=UTF8Length(AData);
  Result:=(LStr>0);
  CountA:=0;
  CountB:=0;
  CountC:=0;
  tmpChrs:='0123456789-'+CHRDOT1+CHRDOT2;
  for i:=1 to LStr do
  begin
     tmpChr:=UTF8Copy(AData,i,1);
     if UTF8Pos(tmpChr,tmpChrs,1)=0 then
     begin
        Result:=False;
        break;
     end;

     if (UTF8Pos('-',AData,1)>1)
        or((UTF8Pos('-',AData,1)>0)and(CountA>1)) then
     begin
        Result:=False;
        break;
     end
     else if ((UTF8Pos('-',AData,1)>0)and(CountA=0)) then
     begin
        inc(CountA);
     end;

     if (UTF8Pos(CHRDOT1,AData,1)=1)
        or((UTF8Pos(CHRDOT1,AData,1)>0)and(CountB>1)) then
     begin
        Result:=False;
        break;
     end
     else if ((UTF8Pos(CHRDOT1,AData,1)>0)and(CountB=0)) then
     begin
        inc(CountB);
     end;

     if (UTF8Pos(CHRDOT2,AData,1)=1)
         or((UTF8Pos(CHRDOT2,AData,1)>0)and(CountC>1)) then
     begin
        Result:=False;
        break;
     end
     else if ((UTF8Pos(CHRDOT2,AData,1)>0)and(CountC=0)) then
     begin
        inc(CountC);
     end;
  end;
end;

function GetCorrectNumber(AData:String):String;
const
  CHRDOT1='.';
  CHRDOT2=',';
begin
  Result:=StringReplace(AData,CHRDOT1,CHRDOT2,[]);
end;

function ArrayCompareText(AText:ShortString;
         const Args : Array of ShortString):Boolean;
var
  i:Integer;
begin
  Result:=False;
  for i:=Low(Args) to High(Args)  do
  begin
    if UTF8CompareStr(Args[i], AText)=0 then
    begin
         Result:=True;
         break;
    end;
  end;
end;

function DeQuoteText(var OutText:String):boolean;
var
  LStr  :Integer;
begin
  LStr  :=UTF8Length(OutText);
  Result:=(LStr>0);

  if UTF8Copy(OutText,1,1)<>QUOTE3 then
  begin
    Result:=False;
  end
  else begin
   if UTF8Copy(OutText,LStr,1)<>QUOTE3 then
   begin
      Result:=False;
   end
   else begin
      OutText:=UTF8Copy(OutText,2,LStr-2);
   end;
  end;
end;

function DeBKTText(var OutText:String):boolean;
var
  i,
  iQuoteOpen,
  LStr       :Integer;
  bFirst,
  bLast,
  bCenter    :boolean;
  sChar      :string;
begin
  LStr         :=UTF8Length(OutText);
  iQuoteOpen   :=0;
  bFirst       :=False;
  bLast        :=False;
  bCenter      :=False;

  for i:=1 to LStr do
  begin
      sChar:=UTF8Copy(OutText, i, 1);
      if (UTF8Pos(sChar,BKT_L)>0) then
      begin
         if i=1 then
            bFirst:=True;
         inc(iQuoteOpen);

      end
      else if (UTF8Pos(sChar,BKT_R)>0) then
      begin
          dec(iQuoteOpen);

          if i=LStr then
            bLast:=True;
      end
      else begin
          if iQuoteOpen=0 then
          begin
             bCenter :=True;
          end;
      end;
  end;
  Result:=False;
  if (not bCenter)and(bFirst and bLast) then
  begin
      OutText:=UTF8Copy(OutText,2,LStr-2);
      Result:=True;
  end;
end;

function GetShieldText(AData:String):String;
begin
  AData :=StringReplace(AData,FORMULA_CHRPARAMSSEPARETER,FORMULA_CHRPARAMSSEPARETER_Alt1,[rfReplaceAll]);
  AData :=StringReplace(AData,QUOTE3,QUOTE1,[rfReplaceAll]);
  Result:=AData;
end;

function GetDeShieldText(AData:String):String;
begin
  AData :=StringReplace(AData,FORMULA_CHRPARAMSSEPARETER_Alt1,FORMULA_CHRPARAMSSEPARETER,[rfReplaceAll]);
  AData :=StringReplace(AData,FORMULA_CHRPARAMSSEPARETER_Alt2,FORMULA_CHRPARAMSSEPARETER,[rfReplaceAll]);
  AData :=StringReplace(AData,QUOTE1,QUOTE3,[rfReplaceAll]);
  AData :=StringReplace(AData,QUOTE2,QUOTE3,[rfReplaceAll]);
  AData :=StringReplace(AData,'&deg;','°',[rfReplaceAll]);
  AData :=StringReplace(AData,'%br10;',#10,[rfReplaceAll]);
  AData :=StringReplace(AData,'%br13;',#13,[rfReplaceAll]);
  Result:=AData;
end;


{ TAssiFormulaParser }

function TAssiFormulaParser.AddFPart: PAFPItem;
var
  c,i:integer;
  tmpItem:PAFPItem;
begin
  c:=Length(FArrayFormulaParts[FParsesIndex]);
  i:=c;
  inc(c);
  SetLength(FArrayFormulaParts[FParsesIndex],c);
  New(tmpItem);
  FArrayFormulaParts[FParsesIndex][i]:=tmpItem;
  Result:=tmpItem;
end;

function TAssiFormulaParser.DoReadString(AFormula: String;
         ALevel:integer=0; AParent:PAFPItem=nil): Boolean;
var
  TextLength,
  iQuoteOpen,
  iQuoteBOpen,
  iCan,
  i,j,c,
  iLevel,
  iCounterChar,
  iCursor,
  iCursorBegin,
  iCursorEnd:integer;

  bError,
  bStart:Boolean;

  sCharLast,
  sChar2,
  sChar,
  sCasheName,
  sCashe:string;

  ParentStatus,
  Status:TAFPStyle;
  FPartItem:PAFPItem;
begin
  sCashe    :='';
  sCharLast :=PChar('');
  Status    :=afpsUnknow;
  iLevel    :=ALevel;

  //AText:=UTF8ToAnsi(AText);

  TextLength       :=UTF8Length(AFormula);
  iCursorBegin     :=1;
  iCursorEnd       :=TextLength;

  iCursor          :=0;
  bError           :=False;
  bStart           :=(TextLength>0);
  ParentStatus     :=afpsUnknow;
  if AParent<>nil then
     ParentStatus  :=AParent^.Style;

  if bStart then
  repeat

    if Status=afpsUnknow then
    begin
       //Определяем что перед нами

       sChar      :=UTF8Copy(AFormula, iCursorBegin, 1);
       sChar      :=UTF8UpperCase(sChar);

       sChar2     :='';
       if iCursorBegin>1 then
          sChar2:=UTF8Copy(AFormula, iCursorBegin-1, 2);

       if (ParentStatus=afpsKeyFun) then
       begin
           Status:=afpsGroupParam;
       end
       else if ArrayCompareText(sChar2,['--','+-','*-','/-','-+','++','*+','/+']) then
       begin
           Status:=afpsNum;
       end
       else if (UTF8Pos(sChar,MATHCHARS)>0)and(iCursorBegin>1) then
       begin
           Status:=afpsKeyMath;
       end
       else if UTF8Pos(sChar,'0123456789-')>0 then
       begin
          Status:=afpsNum;
       end
       else if UTF8Pos(sChar,ABCCHARS)>0 then
       begin
           Status:=afpsKeyFun;
       end
       else if UTF8Pos(sChar,'"')>0 then
       begin
           Status:=afpsText;
       end
       else if UTF8Pos(sChar,'(')>0 then
       begin
           Status:=afpsGroup;
       end
       else begin
         //error
         bError:=True;
       end;
    end;

     sCasheName :='';
     sCashe     :='';
     iCan       :=0;

     if Status=afpsNum then
     begin
        iCounterChar:=0;
        for i:=iCursorBegin to iCursorEnd do
        begin
            sChar:=UTF8Copy(AFormula, i, 1);
            if (UTF8Pos(sChar,'0123456789.,')>0)or((UTF8Pos(sChar,'-')>0)
               and(i=iCursorBegin)) then
            begin
               sCashe :=sCashe+sChar;
               inc(iCan);
               inc(iCounterChar);
            end
            else
            begin
               if (UTF8Pos(sCharLast,'-.,')>0) then
               begin
                  iCan:=0;
               end
               else begin

               end;
               break;
            end;
            sCharLast:=sChar;
        end;
        if iCan>0 then
        begin
           iCursor:=iCursorBegin+iCounterChar;
        end
        else begin
            //error
            bError:=True;
        end;
     end
     else if Status=afpsKeyMath then
     begin
        sCharLast    :='';
        iCounterChar :=0;
        for i:=iCursorBegin to iCursorEnd do
        begin
            sChar:=UTF8Copy(AFormula, i, 1);
            sChar:=UTF8UpperCase(sChar);
            if (UTF8Pos(sChar,MATHCHARS)>0)and(i=iCursorBegin) then
            begin
               sCashe:=sCashe+sChar;
               inc(iCounterChar);
               inc(iCan);
               break;
            end
            else begin
                break;
            end;
            sCharLast:=sChar;
        end;
        if iCan>0 then
        begin
           iCursor:=iCursorBegin+iCounterChar;
        end
        else begin
            //error
            bError:=True;
        end
     end
     else if Status=afpsKeyFun then
     begin
        sCharLast    :='';
        iCounterChar :=0;
        iQuoteOpen   :=0;
        for i:=iCursorBegin to iCursorEnd do
        begin
            sChar:=UTF8Copy(AFormula, i, 1);
            sChar:=UTF8UpperCase(sChar);
            if (UTF8Pos(sChar,ABCCHARS)>0) then
            begin
               sCashe:=sCashe+sChar;
               inc(iCounterChar);
               inc(iCan);
            end
            else begin
                break;
            end;
            sCharLast:=sChar;
        end;

        if iCounterChar>0 then
        begin
            sCasheName :=sCashe;
            sCashe     :='';
            j:=iCursorBegin+iCounterChar;
            for i:=j to iCursorEnd do
            begin
                sChar:=UTF8Copy(AFormula, i, 1);
                if (UTF8Pos(sChar,'(')>0) then
                begin
                   inc(iQuoteOpen);
                   sCashe :=sCashe+sChar;
                   inc(iCounterChar);
                   inc(iCan);
                end
                else if (UTF8Pos(sChar,')')>0) then
                begin
                    dec(iQuoteOpen);
                    sCashe :=sCashe+sChar;
                    inc(iCounterChar);
                    inc(iCan);
                    if iQuoteOpen=0 then
                       break;
                end
                else begin
                   sCashe :=sCashe+sChar;
                   inc(iCounterChar);
                end;
                sCharLast:=sChar;
            end;

            if iQuoteOpen>0 then
               iCan:=0;
        end;

        if iCan>0 then
        begin
           iCursor:=iCursorBegin+iCounterChar;
        end
        else begin
            //error
            bError:=True;
        end
     end
     else if Status=afpsText then
     begin
        sCharLast    :='';
        iCounterChar :=0;
        for i:=iCursorBegin to iCursorEnd do
        begin
            sChar:=UTF8Copy(AFormula, i, 1);
            if (UTF8Pos(sChar,'"')>0)and(i=iCursorBegin) then
            begin
               sCashe :=sCashe+sChar;
               inc(iCounterChar);
               inc(iCan);
            end
            else if (UTF8Pos(sChar,'"')>0)and(sCharLast<>'\') then
            begin
                sCashe :=sCashe+sChar;
                inc(iCounterChar);
                inc(iCan);
                break;
            end
            else begin
               sCashe :=sCashe+sChar;
               inc(iCounterChar);
            end;
            sCharLast:=sChar;
        end;

        if iCan<>2 then
           iCan:=0;

        if iCan>0 then
        begin
           sCashe :=GetDeShieldText(sCashe);
           iCursor:=iCursorBegin+iCounterChar;
        end
        else begin
            //error
            bError:=True;
        end
     end
     else if (Status=afpsGroup) then
     begin
        sCharLast    :='';
        iCounterChar :=0;
        iQuoteOpen   :=0;
        for i:=iCursorBegin to iCursorEnd do
        begin
            sChar:=UTF8Copy(AFormula, i, 1);
            if (UTF8Pos(sChar,'(')>0) then
            begin
               inc(iQuoteOpen);
               sCashe :=sCashe+sChar;
               inc(iCounterChar);
               inc(iCan);
            end
            else if (UTF8Pos(sChar,')')>0) then
            begin
                dec(iQuoteOpen);
                sCashe :=sCashe+sChar;
                inc(iCounterChar);
                inc(iCan);
                if iQuoteOpen=0 then
                   break;
            end
            else begin
               sCashe :=sCashe+sChar;
               inc(iCounterChar);
            end;
            sCharLast:=sChar;
        end;

        if iQuoteOpen>0 then
           iCan:=0;

        if iCan>0 then
        begin
           iCursor:=iCursorBegin+iCounterChar;
        end
        else begin
            //error
            bError:=True;
        end
     end
     else if (Status=afpsGroupParam) then
     begin
        //поиск скобок и кавычек
        sCharLast    :='';
        iCounterChar :=0;
        iQuoteOpen   :=0;
        iQuoteBOpen  :=0;

        for i:=iCursorBegin to iCursorEnd do
        begin
            sChar:=UTF8Copy(AFormula, i, 1);
            if (UTF8Pos(sChar,'(')>0)and(iQuoteBOpen=0) then
            begin
               inc(iQuoteOpen);
            end
            else if (UTF8Pos(sChar,')')>0)and(iQuoteBOpen=0) then
            begin
                dec(iQuoteOpen);
            end
            else if (UTF8Pos(sChar,'"')>0)and(iQuoteBOpen=0) then
            begin
               inc(iQuoteBOpen);
            end
            else if (UTF8Pos(sChar,'"')>0)and(iQuoteBOpen>0) then
            begin
                dec(iQuoteBOpen);
            end
            else if (UTF8Pos(sChar,';')>0)and(iQuoteOpen=0)
               and(iQuoteBOpen=0) then
            begin
               sCashe       :=UTF8Copy(AFormula, iCursorBegin, i-iCursorBegin);
               inc(iCan);
               iCounterChar :=i-iCursorBegin+1;
               break;
            end;
            sCharLast:=sChar;
        end;

        if (iCounterChar=0)and(iCursorBegin<=iCursorEnd) then
        begin
            sCashe       :=UTF8Copy(AFormula, iCursorBegin, iCursorEnd-iCursorBegin+1);
            inc(iCan);
            iCounterChar :=iCursorEnd-iCursorBegin+1;
        end;

        if (iQuoteOpen>0)or(Length(sCashe)=0) then
           iCan:=0;

        if (iCan>0) then
        begin
           iCursor:=iCursorBegin+iCounterChar;
        end
        else begin
            //error
            bError:=True;
        end
     end
     else if Status=afpsUnknow then
     begin
         //error
         bError:=True;
     end;

     if iCan>0 then
     begin
        //создать эл-т данных
        FPartItem                 :=AddFPart;
        FPartItem^.Ignor          :=False;
        FPartItem^.StyleInit      :=Status;
        FPartItem^.Style          :=Status;
        FPartItem^.ContentAsText  :=sCashe;
        FPartItem^.Name           :=sCasheName;
        FPartItem^.Id             :=GetID;
        if AParent<>nil then
           FPartItem^.ParentId    :=AParent^.Id
        else
           FPartItem^.ParentId    :=0;
        FPartItem^.Result         :=null;
        FPartItem^.Content        :=sCashe;
        SetLength(FPartItem^.SubItems, 0);

        if AParent<>nil then
        begin
             c:=Length(AParent^.SubItems);
             j:=c+1;
             SetLength(AParent^.SubItems,j);
             AParent^.SubItems[c]:=FPartItem;
        end;

        if Status=afpsKeyFun then
        begin
           DeBKTText(sCashe);
           bError:=not DoReadString(sCashe,iLevel+1,FPartItem);
        end
        else if Status=afpsGroup then
        begin
           DeBKTText(sCashe);
           bError:=not DoReadString(sCashe,iLevel+1,FPartItem);
        end
        else if Status=afpsGroupParam then
        begin
           DeBKTText(sCashe);
           bError:=not DoReadString(sCashe,iLevel+1,FPartItem);
        end;
     end;

     iCursorBegin:=iCursor;
     if iCursorBegin>iCursorEnd then
        bStart:=False;

     if bError then
        bStart:=False;

     Status:=afpsUnknow;

  until not bStart;//bStart

  if bError and Assigned(FOnFormulaError) then
  begin
       FOnFormulaError(Self, 'Error on parse formula');
  end;

  Result:=not bError;
end;

function TAssiFormulaParser.GroupStyle(AItem :PAFPItem): integer;
var
  i       :integer;
  tmpItem :PAFPItem;

  iGroupParam :integer;
  iGroup      :integer;
  iNum        :integer;
  iKeyFun     :integer;
  iKeyMath    :integer;
  iText       :integer;
begin
  iGroupParam :=0;
  iGroup      :=0;
  iNum        :=0;
  iKeyFun     :=0;
  iKeyMath    :=0;
  iText       :=0;

  for i:=0 to high(AItem^.SubItems) do
  begin
    tmpItem:=AItem^.SubItems[i];
    if (not tmpItem^.Ignor) then
    begin
       if tmpItem^.Style=afpsGroupParam then
       begin
         inc(iGroupParam);
       end
       else if tmpItem^.Style=afpsNum then
       begin
         inc(iNum);
       end
       else if tmpItem^.Style=afpsKeyFun then
       begin
         inc(iKeyFun);
       end
       else if tmpItem^.Style=afpsKeyMath then
       begin
         inc(iKeyMath);
       end
       else if tmpItem^.Style=afpsText then
       begin
          inc(iText);
       end
       else if tmpItem^.Style=afpsGroup then
       begin
          inc(iGroup);
       end;
    end;
  end;

  if (iNum>0)and(iKeyMath>0)
     and(iGroup=0)and(iText=0)and(iGroupParam=0)and(iKeyFun=0) then
  begin
     Result:=1; //математические рассчеты
  end
  else if (iNum=1)and(iKeyMath=0)and(iText=0)
     and(iGroup=0)and(iGroupParam=0)and(iKeyFun=0) then
  begin
     Result:=2; //число
  end
  else if (iNum=0)and(iKeyMath=0)and(iText=1)
     and(iGroup=0)and(iGroupParam=0)and(iKeyFun=0) then
  begin
     Result:=3; //текст
  end
  else if (iNum>0)and(iKeyMath>0)and(iText>0)
     and(iGroup=0)and(iGroupParam=0)and(iKeyFun=0) then
  begin
     Result:=4;//err
  end
  else if (iNum=0)and(iKeyMath=0)and(iText=0)
     and(iGroup=0)and(iGroupParam=0)and(iKeyFun=0) then
  begin
     Result:=5;//clear
  end
  else
  begin
     Result:=0;
  end;
end;

function TAssiFormulaParser.DoCheckSequence(AParentItem :PAFPItem): Boolean;
var
  bError       :boolean;
  i,c          :integer;
  tmpItem2,
  tmpItem1     :PAFPItem;
begin
  bError :=False;
  c      :=high(AParentItem^.SubItems);
  for i:=1 to c do
  begin
      tmpItem1:=AParentItem^.SubItems[i-1];
      tmpItem2:=AParentItem^.SubItems[i];
      if (not tmpItem1^.Ignor)and(not tmpItem2^.Ignor) then
      begin
         if (tmpItem1^.Style in [afpsNum, afpsText])
            and(tmpItem2^.Style in [afpsNum, afpsText]) then
         begin
            bError:=True;
            break;
         end
         else if (tmpItem1^.Style in [afpsKeyMath, afpsText])
            and(tmpItem2^.Style in [afpsKeyMath, afpsText]) then
         begin
            bError:=True;
            break;
         end
         else if (tmpItem1^.Style in [afpsBool, afpsText])
            and(tmpItem2^.Style in [afpsBool, afpsText]) then
         begin
            bError:=True;
            break;
         end
         else if (tmpItem1^.Style in [afpsBool, afpsNum])
            and(tmpItem2^.Style in [afpsBool, afpsNum]) then
         begin
            bError:=True;
            break;
         end
         else if (tmpItem1^.Style in [afpsKeyFun, afpsText])
            and(tmpItem2^.Style in [afpsKeyFun, afpsText]) then
         begin
            bError:=True;
            break;
         end
         else if (tmpItem1^.Style in [afpsKeyFun, afpsNum])
            and(tmpItem2^.Style in [afpsKeyFun, afpsNum]) then
         begin
            bError:=True;
            break;
         end;
      end;
  end;

  if bError and Assigned(FOnFormulaError) then
  begin
       FOnFormulaError(Self, 'Error in param sequence');
  end;

  Result:= not bError;
end;

//Расчет формулы
function TAssiFormulaParser.DoCalcFormula(AParentItem :PAFPItem): Boolean;
var
  vTmpArray            :Array of Variant;
  vProcResult          :Variant;
  bEnd,
  bError               :boolean;
  FloatA,
  FloatB               :Double;
  iStop,
  i,c,k                :integer;
  tmpItemA             :PAFPItem;
  tmpItemB             :PAFPItem;
  tmpItem2,
  tmpItem              :PAFPItem;

  tmpFormulaName       :String;
  TmpList              :TList;
  dfrProc              :TDoFormulaResult;
begin
  bError :=not DoCheckSequence(AParentItem); //Проверка последовательности эл-тов формулы
  bEnd   :=False;
  iStop  :=2000;

  while ((not bError)and(not bEnd)) do
  begin

  i:=GroupStyle(AParentItem);

  if (i=2) then
  begin
    //num
    c:=high(AParentItem^.SubItems);
    TmpList:=TList.Create;

    for i:=0 to c do
    begin
        tmpItem:=AParentItem^.SubItems[i];
        if not tmpItem^.Ignor then
        TmpList.Add(tmpItem);
    end;

    if TmpList.Count<>1 then
    begin
       bError:=True;
    end
    else begin
       tmpItem:=TmpList.Items[0];
       AParentItem^.Result:=tmpItem^.Result;
       AParentItem^.Style:=afpsNum;
       bEnd:=True;
    end;

    TmpList.Free;
  end
  else if (i=3) then
  begin
    //text
    c:=high(AParentItem^.SubItems);
    TmpList:=TList.Create;

    for i:=0 to c do
    begin
        tmpItem:=AParentItem^.SubItems[i];
        if not tmpItem^.Ignor then
        TmpList.Add(tmpItem);
    end;

    if TmpList.Count<>1 then
    begin
       bError:=True;
    end
    else begin
       tmpItem:=TmpList.Items[0];
       AParentItem^.Result:=tmpItem^.Result;
       AParentItem^.Style:=afpsText;
       bEnd:=True;
    end;

    TmpList.Free;

  end
  else if (i=5) then
  begin
     bEnd:=True;
  end
  else if (i=4) then
  begin
     bError:=True;
  end
  else if i=1 then  //Если арифметическое выражение
  begin
     //math
     c:=high(AParentItem^.SubItems);
     TmpList:=TList.Create;
     for i:=0 to c do
     begin
        tmpItem:=AParentItem^.SubItems[i];
        if not tmpItem^.Ignor then
        TmpList.Add(tmpItem);
     end;

     //умножение и деление
     i:=0;
     while (i<TmpList.Count-1)and(TmpList.Count>1) and (TmpList.Count<>0) do
     begin
        tmpItem:=TmpList.Items[i];
        if (tmpItem^.Style=afpsKeyMath)
           and((tmpItem^.Content='*')or(tmpItem^.Content='/')) then
        begin
            if not((0<i)and(i<c)) then
            begin
               bError:=True;
               break;
            end
            else begin
                tmpItemA :=TmpList.Items[i-1];
                tmpItemB :=TmpList.Items[i+1];
                if tmpItemA^.Style<>tmpItemB^.Style then
                begin
                   bError:=True;
                   break;
                end
                else begin
                    FloatA   :=tmpItemA^.Result;
                    FloatB   :=tmpItemB^.Result;
                    if (tmpItem^.Content='*') then
                    begin
                        tmpItem^.Result:=FloatA*FloatB;
                        tmpItem^.Style:=afpsNum;
                    end
                    else if (tmpItem^.Content='/') then
                    begin
                        if FloatB=0 then
                        begin
                          bError:=True;
                          break;
                        end
                        else begin
                          tmpItem^.Result:=FloatA/FloatB;
                          tmpItem^.Style:=afpsNum;
                        end;
                    end;
                    tmpItemA^.Ignor:=True;
                    tmpItemB^.Ignor:=True;
                    TmpList.Remove(tmpItemA);
                    TmpList.Remove(tmpItemB);
                end;
            end;
            i:=0;
        end
        else
            inc(i);
     end;

     //сложение и вычитание

     i:=0;
     while (i<TmpList.Count-1)and(TmpList.Count>1) and (TmpList.Count<>0) do
     begin
        tmpItem:=TmpList.Items[i];
        if (tmpItem^.Style=afpsKeyMath)
           and((tmpItem^.Content='+')or(tmpItem^.Content='-')) then
        begin
            if not((0<i)and(i<c)) then
            begin
               bError:=True;
               break;
            end
            else begin
                tmpItemA :=TmpList.Items[i-1];
                tmpItemB :=TmpList.Items[i+1];
                if tmpItemA^.Style<>tmpItemB^.Style then
                begin
                   bError:=True;
                   break;
                end
                else begin
                    FloatA   :=tmpItemA^.Result;
                    FloatB   :=tmpItemB^.Result;
                    if (tmpItem^.Content='+') then
                    begin
                        tmpItem^.Result:=FloatA+FloatB;
                        tmpItem^.Style:=afpsNum;
                    end
                    else if (tmpItem^.Content='-') then
                    begin
                        tmpItem^.Result:=FloatA-FloatB;
                        tmpItem^.Style:=afpsNum;
                    end;
                    tmpItemA^.Ignor:=True;
                    tmpItemB^.Ignor:=True;
                    TmpList.Remove(tmpItemA);
                    TmpList.Remove(tmpItemB);
                end;
            end;
            i:=0;
        end
        else
            inc(i);
     end;

     if TmpList.Count=1 then
     begin
       tmpItem:=TmpList.Items[0];
       if (tmpItem^.Style=afpsNum) then
       begin
         AParentItem^.Result:=tmpItem^.Result;
         bEnd:=True;
       end;
     end;

     TmpList.Free;

     if bError and Assigned(FOnFormulaError) then
     begin
           FOnFormulaError(Self, 'Error in math function');
     end;
  end
  else
  begin
    //
    c:=high(AParentItem^.SubItems);
    TmpList:=TList.Create;

    for i:=0 to c do
    begin
        tmpItem:=AParentItem^.SubItems[i];
        if not tmpItem^.Ignor then
        TmpList.Add(tmpItem);
    end;

    for i:=0 to TmpList.Count-1 do
    begin
        tmpItem:=TmpList.Items[i];
        AParentItem^.Result:=tmpItem^.Result;

        if tmpItem^.Style=afpsGroup then  //Группа
        begin
           bError:=not DoCalcFormula(tmpItem);
        end
        else if tmpItem^.Style=afpsKeyFun then //Функция
        begin
           tmpFormulaName :=tmpItem^.Name;

           //todo: IfErr() Функция
           //if CompareText('IFERR',tmpFormulaName)=0 then
           //   bIfErr:=True;

           bError         :=not IsReservedName(tmpFormulaName);

           if not bError then
           begin
             k:=length(tmpItem^.SubItems);
             SetLength(vTmpArray,k);
             if (CompareText(tmpFormulaName,'IF')=0) then
             begin
                 //Исключаем выполнение функции Если в двух направлениях.
                 if (k=3) then
                 begin
                     vTmpArray[0] :=null;
                     tmpItem2     :=tmpItem^.SubItems[0];
                     bError       :=not DoCalcFormula(tmpItem2);

                     if bError then
                     begin
                       break;
                     end
                     else begin
                         vTmpArray[0]:=tmpItem2^.Result;

                         if not VarIsBool(vTmpArray[0])then
                         begin
                            bError:=True;
                         end
                         else begin
                            if vTmpArray[0]=True then
                                k:=1
                            else
                                k:=2;
                         end;

                         if not bError then
                         begin
                             vTmpArray[k] :=null;
                             tmpItem2     :=tmpItem^.SubItems[k];
                             bError       :=not DoCalcFormula(tmpItem2);

                             if bError then
                             begin
                               break;
                             end
                             else begin
                               vTmpArray[k]:=tmpItem2^.Result;

                               if k=1 then
                                  k:=2
                               else
                                  k:=1;

                               vTmpArray[k]:='';
                             end;
                          end;
                       end;
                 end
                 else begin
                     bError:=True;
                     if Assigned(FOnFormulaError) then
                     begin
                        FOnFormulaError(Self, format('Wrong param count in "%s" function',[tmpFormulaName]));
                     end;
                 end;
             end
             else begin
                 for k:=0 to high(tmpItem^.SubItems) do
                 begin
                     vTmpArray[k] :=null;
                     tmpItem2     :=tmpItem^.SubItems[k];
                     bError       :=not DoCalcFormula(tmpItem2);

                     if bError then
                     begin
                       break;
                     end
                     else begin
                       vTmpArray[k]:=tmpItem2^.Result;
                     end;
                 end;
             end;
           end
           else begin
                if Assigned(FOnFormulaError) then
                begin
                     FOnFormulaError(Self, format('Unknow "%s" function',[tmpFormulaName]));
                end;
           end;

           if not bError then
           begin
               SetLength(FArrayFormulaParams,length(vTmpArray));
               for k:=0 to high(vTmpArray) do
               begin
                 FArrayFormulaParams[k]:=vTmpArray[k];
               end;
               vProcResult :=null;
               k           :=GetFormulaIndex(tmpFormulaName);
               if k>-1 then
               begin
                  dfrProc:=TDoFormulaResult(FArrayFormulaPointer[k]);
                  dfrProc(Self, vProcResult);
               end;

               if vProcResult<>null then
               begin
                  tmpItem^.Result:=vProcResult;
                  if VarIsNumeric(vProcResult) then
                     tmpItem^.Style:=afpsNum
                  else if VarIsStr(vProcResult) then
                     tmpItem^.Style:=afpsText
                  else if VarIsBool(vProcResult) then
                     tmpItem^.Style:=afpsBool
                  else
                     tmpItem^.Style:=afpsUnknow;
               end
               else begin
                  bError:=True;
                  if Assigned(FOnFormulaError) then
                  begin
                     FOnFormulaError(Self, format('Error in %s function',[tmpFormulaName]));
                  end;
               end;
           end;
           SetLength(FArrayFormulaParams,0);
        end;

        if bError then
           break;
    end;

    TmpList.Free;
  end;

   dec(iStop);  //Защита от зацикливания
   if iStop<1 then
   begin
      bError:=True;
      if Assigned(FOnFormulaError) then
      begin
         FOnFormulaError(Self, 'Выполнение остановлено. Кол-во проходов превысило 2000 раз.');
      end;
   end;

  end;//while

  Result:=not bError;
end;

function TAssiFormulaParser.NewParse: Integer;
var
  i,l:integer;
begin
  l:=Length(FArrayFormulaParts);
  i:=l+1;
  SetLength(FArrayFormulaParts,i);
  Result:=l;
  FParsesIndex:=l;
end;

function TAssiFormulaParser.GetParseCount: Integer;
begin
  Result:=Length(FArrayFormulaParts);
end;

function TAssiFormulaParser.SetCurrentParse(AIndex: Integer): boolean;
begin
  if AIndex<Length(FArrayFormulaParts) then
  begin
    Result:=True;
    FParsesIndex:=AIndex;
  end
  else begin
    Result:=False;
  end;
end;

function TAssiFormulaParser.ClearParses: boolean;
var
  i:integer;
begin
  for i:=0 to high(FArrayFormulaParts) do
  begin
     FParsesIndex:=i;
     ClearFParts;
  end;
  SetLength(FArrayFormulaParts,0);
  NewParse;
end;

procedure TAssiFormulaParser.ClearFParts;
var
  i:integer;
  tmpItem:PAFPItem;
begin
  for i:=0 to high(FArrayFormulaParts[FParsesIndex]) do
  begin
     tmpItem:=FArrayFormulaParts[FParsesIndex][i];
     Dispose(tmpItem);
  end;
  SetLength(FArrayFormulaParts[FParsesIndex],0);
end;

function TAssiFormulaParser.GetParentStyle(AID: integer): TAFPStyle;
var
  i        :integer;
  tmpItem  :PAFPItem;
begin
  Result:=afpsUnknow;
  for i:=0 to high(FArrayFormulaParts[FParsesIndex]) do
  begin
      tmpItem:=FArrayFormulaParts[FParsesIndex][i];
      if (tmpItem^.ID=AID) then
      begin
          Result:=tmpItem^.Style;
      end;
  end;
end;

function TAssiFormulaParser.GetID: integer;
begin
  inc(FIDCounter);
  Result:=FIDCounter;
end;

function TAssiFormulaParser.IsReservedName(AName: String): Boolean;
begin
  Result:=GetFormulaIndex(AName)>-1;
end;

function TAssiFormulaParser.GetFormulaIndex(AName: String): Integer;
var
  i:integer;
begin
   Result:=-1;
   AName:=UTF8UpperCase(AName);
   for i:=0 to high(FArraySynonymNameAlternative) do
   begin
       if CompareText(FArraySynonymNameAlternative[i], AName)=0 then
       begin
          AName:=FArraySynonymNameOriginal[i];
          break;
       end;
   end;
   for i:=0 to high(FArrayFormulaName) do
   begin
       if CompareText(FArrayFormulaName[i], AName)=0 then
       begin
          Result:=i;
          break;
       end;
   end;
end;

function TAssiFormulaParser.Parse(AFormula: String): Boolean;
begin
  ClearFParts;
  FIDCounter:=0;
  FResult:=null;
  Result:=DoReadString(AFormula);
end;

function TAssiFormulaParser.Execute(AFormula: String): Boolean;
begin
  Result:=Parse(AFormula);
  if Result then
  begin
     Result:=Perform;
  end;
end;

function TAssiFormulaParser.Perform: Boolean;
var
  i,c,j   :integer;
  Item    :TAFPItem;
  tmpItem :PAFPItem;
  tmpFloat:Double;
  tmpStr  :string;
begin

  //Перед расчетом сброс предыдущих расчетов по этой формуле
  for i:=0 to high(FArrayFormulaParts[FParsesIndex]) do
  begin
    tmpItem:=FArrayFormulaParts[FParsesIndex][i];
    tmpItem^.Ignor          :=False;
    tmpItem^.Style          :=tmpItem^.StyleInit;
    tmpItem^.Result         :=null;
  end;

  //Расчет
  SetLength(Item.SubItems,0);
  for i:=0 to high(FArrayFormulaParts[FParsesIndex]) do
  begin
    tmpItem:=FArrayFormulaParts[FParsesIndex][i];
    if (tmpItem^.ParentID=0) then
    begin
        c:=Length(Item.SubItems);
        j:=c+1;
        SetLength(Item.SubItems,j);
        Item.SubItems[c]:=tmpItem;
    end;

    //Выставляем результат для простых объектов
    if tmpItem^.Style=afpsNum then
    begin
       tmpFloat:=StrToFloat(tmpItem^.Content,DefaultFormatSettings);
       tmpItem^.Result:=tmpFloat;
    end
    else if tmpItem^.Style=afpsText then
    begin
       tmpStr:=VarToStr(tmpItem^.Content);
       DeQuoteText(tmpStr);
       tmpItem^.Result:=tmpStr;
    end
    else begin
       tmpItem^.Result:=null;
    end;

  end;

  Result:=DoCalcFormula(@Item);
  if Result then
     FResult:=Item.Result;
end;

procedure TAssiFormulaParser.UploadToTree(ATree: TTreeView;
         ParentNode:TTreeNode=nil; AIDParent:integer=0);
var
  i :integer;
  tmpItem  :PAFPItem;
  Node     :TTreeNode;
begin
  for i:=0 to high(FArrayFormulaParts[FParsesIndex]) do
  begin
      tmpItem:=FArrayFormulaParts[FParsesIndex][i];
      if (tmpItem^.ParentID=AIDParent) then
      begin
          Node:=ATree.Items.AddChild(ParentNode,format('[%d] %s %s',[tmpItem^.ID,tmpItem^.Name, tmpItem^.ContentAsText]));
          UploadToTree(ATree,Node,tmpItem^.ID);
      end;
  end;
end;

function TAssiFormulaParser.GetFormulaParams(AIndex: integer): Variant;
begin
  Result:=FArrayFormulaParams[AIndex];
end;

function TAssiFormulaParser.GetFormulaParamCount: integer;
begin
  Result:=Length(FArrayFormulaParams);
end;

function TAssiFormulaParser.GetResult: Variant;
begin
  Result:=FResult;
end;

function TAssiFormulaParser.GetResultAsText: String;
begin
  if FResult<>null then
     Result:=VarToStr(FResult)
  else
     Result:='';
end;

function TAssiFormulaParser.GetFormulaCount: Integer;
begin
  Result:=Length(FArrayFormulaName);
end;

procedure TAssiFormulaParser.FormulaSynonymAdd(AOriginalName,
         ASynonymName:String);
var
   i:integer;
begin
   i:=Length(FArraySynonymNameOriginal);
   inc(i);
   SetLength(FArraySynonymNameOriginal,i);
   SetLength(FArraySynonymNameAlternative,i);
   dec(i);
   FArraySynonymNameOriginal[i]    :=UTF8UpperCase(AOriginalName);
   FArraySynonymNameAlternative[i] :=UTF8UpperCase(ASynonymName);
end;

procedure TAssiFormulaParser.FormulaAddDefault;
begin
  SetDefaultFormula(Self);
end;

procedure TAssiFormulaParser.FormulaAdd(AName:String; AFunc:Pointer);
var
   i:integer;
begin
   i:=Length(FArrayFormulaName);
   inc(i);
   SetLength(FArrayFormulaName,i);
   SetLength(FArrayFormulaPointer,i);
   dec(i);
   FArrayFormulaName[i]    :=UTF8UpperCase(AName);
   FArrayFormulaPointer[i] :=AFunc;
end;

procedure TAssiFormulaParser.FormulaInitNew(ANewCount: Integer);
var
   i:integer;
begin
   i:=Length(FArrayFormulaName);
   i:=i+ANewCount;
   SetLength(FArrayFormulaName,i);
   SetLength(FArrayFormulaPointer,i);
end;

procedure TAssiFormulaParser.FormulaInsert(Index: integer; AName: String;
  AFunc: Pointer);
begin
  FArrayFormulaName[Index]    :=UTF8UpperCase(AName);
  FArrayFormulaPointer[Index] :=AFunc;
end;

constructor TAssiFormulaParser.Create;
begin
  inherited Create;
  FOnFormulaError:=nil;
  FIDCounter:=0;
  NewParse;
  SetLength(FArrayFormulaParts[FParsesIndex],0);
  SetLength(FArrayFormulaName,0);
  SetLength(FArrayFormulaPointer,0);
  SetLength(FArrayFormulaParams,0);
  SetLength(FArraySynonymNameOriginal,0);
  SetLength(FArraySynonymNameAlternative,0);
  FormulaAddDefault();
end;

destructor TAssiFormulaParser.Destroy;
begin
  ClearParses;
  inherited Destroy;
end;

end.
