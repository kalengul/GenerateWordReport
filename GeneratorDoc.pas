unit GeneratorDoc;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Menus, Mask, ComObj,padegFIO,PageDoc, Spin, Grids,
  ExtCtrls;

type
  TFMain = class(TForm)
    btnExit: TButton;
    BtnAddData: TButton;
    BtnEndAddData: TButton;
    LblNameBD: TLabel;
    EdtNameBD: TEdit;
    BtOpenBd: TButton;
    dlgOpen1: TOpenDialog;
    Label1: TLabel;
    Pn1: TPanel;
    Pn2: TPanel;
    Pn3: TPanel;
    SgElementBase: TStringGrid;
    Pn4: TPanel;
    Pn5: TPanel;
    Pn6: TPanel;
    RgFormulation: TRadioGroup;
    Memo1: TMemo;
    BtAddPole: TButton;
    BtAddAllFormulation: TButton;
    Label2: TLabel;
    Bttest: TButton;
    Od1: TOpenDialog;
    procedure btnExitClick(Sender: TObject);
    procedure BtOpenBdClick(Sender: TObject);
    procedure BtnAddDataClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BtnEndAddDataClick(Sender: TObject);
    procedure SeNomBDRowChange(Sender: TObject);
    procedure RgFormulationClick(Sender: TObject);
    procedure BtAddPoleClick(Sender: TObject);
    procedure BtAddAllFormulationClick(Sender: TObject);
    procedure BttestClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

Function ChangeStringAboutFormular(ChangeString:String):string;
Function GoCellsInStringGrid(WordString:String):string;
function GoFormulation(WordString:String):string;

var
  FMain: TFMain;
  Excel: Variant;
  Decl : Variant;
  NomStr:LongWord;
  NomColExcel:LongWord;
  BdConnect:Boolean;

implementation

{$R *.dfm}

procedure TFMain.btnExitClick(Sender: TObject);
begin
FMain.Close;
end;

Function PreobrazovanieKovichek(st:string):string;
begin
while pos(' "',st)<>0 do
  St[pos(' "',st)+1]:='«';
while pos('("',st)<>0 do
  St[pos('("',st)+1]:='«';
while pos('$"',st)<>0 do
  St[pos('$"',st)+1]:='«';
while pos('+"',st)<>0 do
  St[pos('+"',st)+1]:='«';
while pos('"',st)<>0 do
  St[pos('"',st)]:='»';
Result:=st;
end;

procedure TFMain.BtOpenBdClick(Sender: TObject);
var
  ExcelStr,ExcelFStr:string;
begin
if dlgOpen1.Execute then
  begin
  Excel.Workbooks.Open(dlgOpen1.FileName);
  NomStr:= Excel.Cells[1,2];
  EdtNameBD.Text:= dlgOpen1.FileName;
  BdConnect:=True;
  SeNomBDRow.Value:=NomStr;
  SeNomBDRow.MaxValue:=NomStr;
  ShowMessage('Подключена база данных '+dlgOpen1.FileName);
  NomColExcel:=1;
  ExcelStr:=Excel.Cells[3,NomColExcel];
  While ExcelStr<>'' do
    begin
    ExcelFStr:=Excel.Cells[2,NomColExcel];

    SgElementBase.RowCount:=NomColExcel+1;
    SgElementBase.Cells[0,NomColExcel]:=ExcelStr;
    if ExcelFStr<>'' then
      SgElementBase.Cells[1,NomColExcel]:=ExcelFStr;
    Inc(NomColExcel);
    ExcelStr:=Excel.Cells[3,NomColExcel];
    end;
  Dec(NomColExcel);
  SeNomField.MaxValue:=NomColExcel;
  end;
end;

Procedure LoadStringInDB(NomStr:LongWord);
var
  ColExcelNi:LongWord;
  ExcelStr:string;
begin
with FMain do
  begin
  ColExcelNi:=1;
  While ColExcelNi<=NomColExcel do
    begin
    if RgFormulation.ItemIndex=1 then
      ExcelStr:=Excel.Cells[2,ColExcelNi];
    if (RgFormulation.ItemIndex=0) or (ExcelStr='') then
      ExcelStr:=Excel.Cells[NomStr,ColExcelNi];

    SgElementBase.Cells[1,ColExcelNi]:=ExcelStr;
    inc(ColExcelNi);
    end;
  end;
end;

Function GoCellsInStringGrid(WordString:String):string;
var
  NomStringZam:LongWord;
  StartStr,SearchStr,EndStr:string;
begin
While (WordString<>'') and ((Pos('"',WordString)<>0) or (Pos('«',WordString)<>0)) do
  begin
  If (Pos('"',WordString)<>0) then
    begin
    StartStr:=Copy(WordString,1,Pos('"',WordString)-1);
    Delete(WordString,1,Pos('"',WordString));
    SearchStr:=Copy(WordString,1,Pos('"',WordString)-1);
    Delete(WordString,1,Pos('"',WordString));
    EndStr:=WordString
    end
  else
  If (Pos('«',WordString)<>0) then
    begin
    StartStr:=Copy(WordString,1,Pos('«',WordString)-1);
    Delete(WordString,1,Pos('«',WordString));
    SearchStr:=Copy(WordString,1,Pos('»',WordString)-1);
    Delete(WordString,1,Pos('»',WordString));
    EndStr:=WordString
    end;

  NomStringZam:=1;
  while (NomStringZam<=FMain.SgElementBase.RowCount) and (SearchStr<>FMain.SgElementBase.Cells[0,NomStringZam]) do
    NomStringZam:=NomStringZam+1;

  if NomStringZam<=FMain.SgElementBase.RowCount then
    SearchStr:= FMain.SgElementBase.Cells[1,NomStringZam];

  WordString:=StartStr+SearchStr+EndStr;
  end;
Result:=WordString;
end;

function GoFormulation(WordString:String):string;
var
  LenStr:LongWord;
  FunctNameStr,FileString,Par0Str,Par1Str,Par2Str,Par3Str,TempStr:string;
begin
if (WordString<>'') and (WordString[1]='f') then
  begin
  Delete(WordString,1,1);
  FunctNameStr:=Copy(WordString,1,Pos('(',WordString)-1);
  Delete(WordString,1,Pos('(',WordString));
  LenStr:= Length(WordString);
  SetLength(WordString,LenStr-1);

  WordString:=GoCellsInStringGrid(WordString);

  If FunctNameStr='ИП' then
    WordString:=GetFIOPadegFSAS(WordString,1)
  else If FunctNameStr='РП' then
    WordString:=GetFIOPadegFSAS(WordString,2)
  else If FunctNameStr='ДП' then
    WordString:=GetFIOPadegFSAS(WordString,3)
  else If FunctNameStr='ВП' then
    WordString:=GetFIOPadegFSAS(WordString,4)
  else If FunctNameStr='ТП' then
    WordString:=GetFIOPadegFSAS(WordString,5)
  else If FunctNameStr='ПП' then
    WordString:=GetFIOPadegFSAS(WordString,6)
  else If FunctNameStr='ИП1' then
    WordString:=GetAppointmentPadeg(WordString,1)
  else If FunctNameStr='РП1' then
    WordString:=GetAppointmentPadeg(WordString,2)
  else If FunctNameStr='ДП1' then
    WordString:=GetAppointmentPadeg(WordString,3)
  else If FunctNameStr='ВП1' then
    WordString:=GetAppointmentPadeg(WordString,4)
  else If FunctNameStr='ТП1' then
    WordString:=GetAppointmentPadeg(WordString,5)
  else If FunctNameStr='ПП1' then
    WordString:=GetAppointmentPadeg(WordString,6)
  else If FunctNameStr='ИПОр' then
    WordString:=GetOfficePadeg(WordString,1)
  else If FunctNameStr='РПОр' then
    WordString:=GetOfficePadeg(WordString,2)
  else If FunctNameStr='ДПОр' then
    WordString:=GetOfficePadeg(WordString,3)
  else If FunctNameStr='ВПОр' then
    WordString:=GetOfficePadeg(WordString,4)
  else If FunctNameStr='ТПОр' then
    WordString:=GetOfficePadeg(WordString,5)
  else If FunctNameStr='ППОр' then
    WordString:=GetOfficePadeg(WordString,6)
  else If FunctNameStr='ИПФИО' then
    begin
    Par0Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    Par1Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    WordString:=GetFIOPadegAS(Par0Str,Par1Str,WordString,1)
    end
  else If FunctNameStr='РПФИО' then
    begin
    Par0Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    Par1Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    WordString:=GetFIOPadegAS(Par0Str,Par1Str,WordString,2)
    end
  else If FunctNameStr='ДПФИО' then
    begin
    Par0Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    Par1Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    WordString:=GetFIOPadegAS(Par0Str,Par1Str,WordString,3)
    end
  else If FunctNameStr='ВПФИО' then
    begin
    Par0Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    Par1Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    WordString:=GetFIOPadegAS(Par0Str,Par1Str,WordString,4)
    end
  else If FunctNameStr='ТПФИО' then
    begin
    Par0Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    Par1Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    WordString:=GetFIOPadegAS(Par0Str,Par1Str,WordString,5)
    end
  else If FunctNameStr='ППФИО' then
    begin
    Par0Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    Par1Str:=Copy(WordString,1,Pos(' ',WordString)-1);
    Delete(WordString,1,Pos(' ',WordString));
    WordString:=GetFIOPadegAS(Par0Str,Par1Str,WordString,6)
    end
  else If FunctNameStr='Ин' then
    WordString:=WordString[1]
  else If FunctNameStr='Пол' then
    begin
    If ((WordString[Length(WordString)]='ч')or (WordString[Length(WordString)-1]='и')) then
      WordString:='М'
    else
      WordString:='Ж';
    end
  else If FunctNameStr='ВордПоиск' then
    begin
    Par0Str:=Copy(WordString,1,Pos(',',WordString)-1);
    Par0Str:=GoCellsInStringGrid(Par0Str);
    Delete(WordString,1,Pos(',',WordString));
    Par1Str:=Copy(WordString,1,Pos(',',WordString)-1);
    Par1Str:=GoCellsInStringGrid(Par1Str);
    Delete(WordString,1,Pos(',',WordString));
    Par2Str:=WordString;
    Word:=CreateOleObject('Word.Application');
    Word.Documents.Open(GetCurrentDir+'\'+Par0Str);
    FileString:=Word.ActiveDocument.Range.Text;
    if Pos(Par1Str,FileString)<>0 then
      WordString:=Copy(FileString,Pos(Par1Str,FileString)+Length(Par1Str),StrToInt(Par2Str))
    else
      WordString:='';
    Word.Documents.Close;
    Word.Quit;
    Word:=UnAssigned;
    end
  else If FunctNameStr='ВордПоискСмещ' then  //$fВордПоискСмещ("Фамилия"+1.docx,Серия/Номер,2,11)
    begin
    Par0Str:=Copy(WordString,1,Pos(',',WordString)-1);
    Par0Str:=GoCellsInStringGrid(Par0Str);
    Delete(WordString,1,Pos(',',WordString));
    Par1Str:=Copy(WordString,1,Pos(',',WordString)-1);
    Par1Str:=GoCellsInStringGrid(Par1Str);
    Delete(WordString,1,Pos(',',WordString));
    Par2Str:=Copy(WordString,1,Pos(',',WordString)-1);
    Delete(WordString,1,Pos(',',WordString));
    Par3Str:=WordString;
    Word:=CreateOleObject('Word.Application');
    Word.Documents.Open(GetCurrentDir+'\'+Par0Str);
    FileString:=Word.ActiveDocument.Range.Text;
    if Pos(Par1Str,FileString)<>0 then
      WordString:=Copy(FileString,Pos(Par1Str,FileString)+Length(Par1Str)+StrToInt(Par2Str),StrToInt(Par3Str))
    else
      WordString:='';
    Word.Documents.Close;
    Word.Quit;
    Word:=UnAssigned;
    end
  else If FunctNameStr='ВордПоискСимв' then
    begin
    Par0Str:=Copy(WordString,1,Pos(',',WordString)-1);
    Par0Str:=GoCellsInStringGrid(Par0Str);
    Delete(WordString,1,Pos(',',WordString));
    Par1Str:=Copy(WordString,1,Pos(',',WordString)-1);
    Par1Str:=GoCellsInStringGrid(Par1Str);
    Delete(WordString,1,Pos(',',WordString));
    Par2Str:=Copy(WordString,1,Pos(',',WordString)-1);
    Delete(WordString,1,Pos(',',WordString));
    Par3Str:=WordString;
    Par3Str:=GoCellsInStringGrid(Par3Str);
    Word:=CreateOleObject('Word.Application');
    Word.Documents.Open(GetCurrentDir+'\'+Par0Str);
    FileString:=Word.ActiveDocument.Range.Text;
    if Pos(Par1Str,FileString)<>0 then
      begin
      Delete(FileString,1,Pos(Par1Str,FileString)+Length(Par1Str));
      WordString:=Copy(FileString,StrToInt(Par2Str),Pos(Par3Str,FileString))
      end
    else
      WordString:='';
    Word.Documents.Close;
    Word.Quit;
    Word:=UnAssigned;
    end
  end;
Result:=WordString;
end;

Function ChangeStringAboutFormular(ChangeString:String):string;
var
  BasicString,WordString,FormularString,NewString:string;
  SgString:string;
  NomEnd:LongWord;
begin
  Delete(ChangeString,1,1);
  BasicString:=ChangeString;
  NewString:='';
{  If (BasicString[1]='f') and (Pos('Ворд',BasicString)<>0) and (Pos('(',BasicString)<>0) then
    begin
    FormularString:=Copy(BasicString,1,Pos('(',BasicString)-1);
    Delete(BasicString,1,Pos('(',BasicString));
    end
  else
    FormularString:='';     }
  If Pos('+',BasicString)<>0 then
    NomEnd:=0
  else
    NomEnd:=1;
  while NomEnd<>2 do
    begin
    If NomEnd=1 then
      WordString:=BasicString
    else
      begin
      WordString:=Copy(BasicString,1,Pos('+',BasicString)-1);
      Delete(BasicString,1,Pos('+',BasicString));
      end;

    WordString:=GoCellsInStringGrid(WordString);
    WordString:=GoFormulation(WordString);

    If (WordString<>'') and(WordString[1]='$') then
      WordString:=ChangeStringAboutFormular(WordString);

    NewString:=NewString+WordString;

    If NomEnd=1 then
      NomEnd:=2
    else
    If Pos('+',BasicString)<>0 then
       NomEnd:=0
    else
       NomEnd:=1;
    end;
{  if FormularString<>'' then
    begin
    WordString:=FormularString+BasicString;
    NewString:=GoFormulation(WordString);
    end;         }
  Result:=NewString;
end;

procedure TFMain.BtnAddDataClick(Sender: TObject);
var
  KolStrExcel:LongWord;
  StringToExcel,NameToExcel:string;
  ColExcelNi:LongWord;
begin
If BdConnect then
  begin
  NomStr:= SeNomBDRow.Value;

  ColExcelNi:=1;
  While ColExcelNi<=NomColExcel do
    begin
    StringToExcel:=SgElementBase.Cells[1,ColExcelNi];
    NameToExcel:=SgElementBase.Cells[0,ColExcelNi];
    If StringToExcel[1]='$' then
      begin
      Excel.Cells[2,ColExcelNi]:=StringToExcel;
      StringToExcel:=ChangeStringAboutFormular(StringToExcel);
      end;
    Excel.Cells[3,ColExcelNi]:=NameToExcel;
    Excel.Cells[NomStr,ColExcelNi]:=StringToExcel;
    inc(ColExcelNi);
    end;

  KolStrExcel:=Excel.Cells[1,2];
  IF NomStr=KolStrExcel then
    begin
    NomStr:=NomStr+1;
    Excel.Cells[1,2]:=NomStr;
    SeNomBDRow.Value:=NomStr;
    SeNomBDRow.MaxValue:=NomStr;
    end;
  ShowMessage('Данные успешно добавлены в БД '+dlgOpen1.FileName);
  LoadStringInDB(NomStr);
  //Excel.ActiveWorkbook.SaveAs(dlgOpen1.FileName);
  end
else
  ShowMessage('База не подключена');
end;

procedure TFMain.SeNomBDRowChange(Sender: TObject);
begin
If BdConnect then
  begin
  NomStr:= SeNomBDRow.Value;
  LoadStringInDB(NomStr);
  end
else
  ShowMessage('База не подключена');
end;

procedure TFMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Excel.Workbooks.close;
//Excel.close;
end;

procedure TFMain.FormCreate(Sender: TObject);
begin
Excel := CreateOleObject('Excel.Application');
BdConnect:=False;
ShablonLoad:=False;
NomStr:=4;
SgElementBase.ColWidths[0]:=200;
SgElementBase.ColWidths[1]:=750;
SgElementBase.Cells[0,0]:='Параметр';
SgElementBase.Cells[1,0]:='Данные';
end;

procedure TFMain.BtnEndAddDataClick(Sender: TObject);
begin
If BdConnect then
  begin
  NomColumsExcel:=SeNomField.Value;
  FDocuments.ShowModal;
  end
else
  ShowMessage('База не подключена');
end;

procedure TFMain.RgFormulationClick(Sender: TObject);
var
  NomRow:LongWord;
  StringToExcel:string;
begin
For NomRow:=1 to SgElementBase.RowCount-1 do
  begin
  StringToExcel:=SgElementBase.Cells[1,NomRow];
  If StringToExcel<>'' then
    begin
    If (RgFormulation.ItemIndex=0) and (StringToExcel[1]='$') then
      begin
      Excel.Cells[2,NomRow]:=StringToExcel;
      StringToExcel:=ChangeStringAboutFormular(StringToExcel);
      end;
    If RgFormulation.ItemIndex=1 then
      begin
      StringToExcel:=Excel.Cells[2,NomRow];
      end;
    end;
  If StringToExcel<>'' then
    begin
    StringToExcel:= PreobrazovanieKovichek(StringToExcel);
    SgElementBase.Cells[1,NomRow]:=StringToExcel;
    end;
  end;
//LoadStringInDB(NomStr);
end;

procedure TFMain.BtAddPoleClick(Sender: TObject);
var
  Row:LongWord;
begin
Row:=SgElementBase.RowCount;
SgElementBase.RowCount:=Row+1;
NomColExcel:=NomColExcel+1;
SeNomField.MaxValue:=NomColExcel;
end;

procedure TFMain.BtAddAllFormulationClick(Sender: TObject);
var
  NomRow,NomCol:LongWord;
  StringToExcel,StringF:string;
begin
If BdConnect then
begin
For NomRow:=4 to NomStr do
  begin
  LoadStringInDB(NomRow);
  For NomCol:=1 to NomColExcel do
    begin
    StringToExcel:=Excel.Cells[NomRow,NomCol];
    StringF:= Excel.Cells[2,NomCol];
    If (StringF<>'') and (StringF[1]='$') then
        begin

        StringToExcel:=ChangeStringAboutFormular(StringF);
        end;
    If StringToExcel<>'' then
      begin
      StringToExcel:= PreobrazovanieKovichek(StringToExcel);
      Excel.Cells[NomRow,NomCol]:=StringToExcel;
      end;
    end;
  end;
  ShowMessage('Формулы применены');
end
else
  ShowMessage('База не подключена');
end;

procedure TFMain.BttestClick(Sender: TObject);
var
  st,st1:string;
begin
If Od1.Execute then
  begin
  Word:=CreateOleObject('Word.Application');
  Word.Documents.Open(Od1.FileName);
  st:=Word.ActiveDocument.Range.Text;
  st1:='Серия/Номер';
  if Pos(st1,st)<>0 then
    ShowMessage(Copy(st,Pos(st1,st)+Length(st1)+2,11));

  Word.Documents.Close;
  end;
end;

end.
