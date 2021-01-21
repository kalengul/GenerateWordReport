unit PageDoc;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids,ComObj;

type
  TFDocuments = class(TForm)
    Pn1: TPanel;
    Pn2: TPanel;
    BtLOadDocDocument: TButton;
    Od1: TOpenDialog;
    LbAllPeople: TListBox;
    MeSelectPeople: TMemo;
    BtAddSelectPeople: TButton;
    SgSetting: TStringGrid;
    PnP: TPanel;
    BtGoFile: TButton;
    BtAddSelectPeopleAll: TButton;
    procedure BtLOadDocDocumentClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure BtAddSelectPeopleClick(Sender: TObject);
    procedure BtGoFileClick(Sender: TObject);
    procedure BtAddSelectPeopleAllClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FDocuments: TFDocuments;
  NameShablonFile,TypeShablonFile:string;
  Word: variant;
  ShablonLoad:Boolean;
  NomColumsExcel:LongWord;

implementation

{$R *.dfm}

Uses GeneratorDoc;

procedure TFDocuments.BtLOadDocDocumentClick(Sender: TObject);
begin
If Od1.Execute then
  begin
  NameShablonFile:=Od1.FileName;
  //Удалить расширение файла
  If Copy(NameShablonFile,Length(NameShablonFile)-3,4)='.doc' then
    begin
    Delete(NameShablonFile,Length(NameShablonFile)-3,4);
    TypeShablonFile:='.doc';
    ShablonLoad:=True;
    ShowMessage('Загружен шаблон '+Od1.FileName);
    end
  else if Copy(NameShablonFile,Length(NameShablonFile)-4,5)='.docx' then
    begin
    Delete(NameShablonFile,Length(NameShablonFile)-4,5);
    TypeShablonFile:='.docx';
    ShablonLoad:=True;
    ShowMessage('Загружен шаблон '+Od1.FileName);
    end
  else
    begin
    ShablonLoad:=False;
    ShowMessage('Файл '+Od1.FileName+' не является файлом MSWORD');
    end;
  end;
end;

procedure TFDocuments.FormActivate(Sender: TObject);
var
  i:LongWord;
  NomStr:LongWord;
  st:string;
begin
NomStr:= Excel.Cells[1,2];
i:=4;
LbAllPeople.Clear;
While NomStr>i do
  begin
  st:=Excel.Cells[i,NomColumsExcel];
  St:=IntToStr(i)+'$'+st;
  LbAllPeople.Items.Add(st);
  Inc(i);
  end;

SgSetting.Cells[0,0]:='Excel position';
SgSetting.Cells[1,0]:='Word point';
SgSetting.RowCount:=NomColExcel*2+1;
for i:=1 to NomColExcel do
  begin
  st:=Excel.Cells[3,i];
  SgSetting.Cells[0,i]:=st;
  SgSetting.Cells[1,i]:='%$'+st+'$%';
  end;
for i:=1 to NomColExcel do
  begin
  st:=Excel.Cells[3,i];
  SgSetting.Cells[0,NomColExcel+i]:='Список '+st;
  SgSetting.Cells[1,NomColExcel+i]:='%Спис$'+st+'$%';
  end;
end;

procedure TFDocuments.BtAddSelectPeopleClick(Sender: TObject);
begin
If LbAllPeople.ItemIndex<>-1 then
MeSelectPeople.Lines.Add(LbAllPeople.Items[LbAllPeople.ItemIndex]);
end;

function FindAndReplace(const FindText,ReplaceText:string):boolean;
  const wdReplaceAll = 2;
begin
  Word.Selection.Find.MatchSoundsLike := False;
  Word.Selection.Find.MatchAllWordForms := False;
  Word.Selection.Find.MatchWholeWord := False;
  Word.Selection.Find.Format := False;
  Word.Selection.Find.Forward := True;
  Word.Selection.Find.ClearFormatting;
  Word.Selection.Find.Text:=FindText;
  Word.Selection.Find.Replacement.Text:=ReplaceText;
  FindAndReplace:=Word.Selection.Find.Execute(Replace:=wdReplaceAll);
end;

function FindAndReplaceRetry(const FindText,ReplaceText:string):boolean;
  const wdReplaceAll = 2;
begin
  Word.Selection.Find.MatchSoundsLike := False;
  Word.Selection.Find.MatchAllWordForms := False;
  Word.Selection.Find.MatchWholeWord := False;
  Word.Selection.Find.Format := False;
  Word.Selection.Find.Forward := True;
  Word.Selection.Find.ClearFormatting;
  Word.Selection.Find.Text:=FindText;
  Word.Selection.Find.Replacement.Text:=ReplaceText+Chr(13)+FindText;
  FindAndReplaceRetry:=Word.Selection.Find.Execute(Replace:=wdReplaceAll);
end;

Procedure DeleteAllSymbolStr(var St:string; St1:string);
var
  nst1:LongWord;
begin
  nst1:=Length(st1);
  While Pos(St1,st)<>0 do
    Delete(St,Pos(St1,st)-1,nst1);
end;

procedure TFDocuments.BtGoFileClick(Sender: TObject);
var
  NomStrPeople:LongWord;
  StrName,st1,StrTable:string;
  StrFirstFile,StrSecondFile:string;
  NomStrExcel,NomStrExcelWhile,NomPosTabled:LongWord;
begin
If ShablonLoad then
begin
NomStrPeople:=0;
StrFirstFile:=NameShablonFile+TypeShablonFile;
Word:=CreateOleObject('Word.Application');
While NomStrPeople<=MeSelectPeople.Lines.Count-1 do
  begin
  StrName:=MeSelectPeople.Lines[NomStrPeople];
  st1:=Copy(StrName,1,Pos('$',StrName)-1);
  NomStrExcel:=StrToInt(st1);
  Delete(StrName,1,Pos('$',StrName));
  DeleteAllSymbolStr(StrName,'~');
  DeleteAllSymbolStr(StrName,'#');
  DeleteAllSymbolStr(StrName,'%');
  DeleteAllSymbolStr(StrName,'&');
  DeleteAllSymbolStr(StrName,'*');
  DeleteAllSymbolStr(StrName,'{');
  DeleteAllSymbolStr(StrName,'}');
  DeleteAllSymbolStr(StrName,'\');
  DeleteAllSymbolStr(StrName,'/');
  DeleteAllSymbolStr(StrName,':');
  DeleteAllSymbolStr(StrName,'<');
  DeleteAllSymbolStr(StrName,'>');
  DeleteAllSymbolStr(StrName,'?');
  DeleteAllSymbolStr(StrName,'+');
  DeleteAllSymbolStr(StrName,'|');
  DeleteAllSymbolStr(StrName,'"');
  StrSecondFile:=NameShablonFile+' '+StrName+TypeShablonFile;
  CopyFile(PAnsiChar(StrFirstFile),PAnsiChar(StrSecondFile),True);
  Word.Documents.Open(StrSecondFile);
  NomPosTabled:=1;
  while NomPosTabled<=SgSetting.RowCount-1 do
    begin
    StrTable:=SgSetting.Cells[1,NomPosTabled];
    If StrTable[2]='$' then
      begin
      st1:=Excel.Cells[NomStrExcel,NomPosTabled];
      FindAndReplace(StrTable,st1);
      end;
{    else if StrTable[2]='С' then
      begin
      NomStrExcelWhile:=3;
      while NomStrExcelWhile<FMain.SeNomBDRow.MaxValue do
        begin
        st1:=Excel.Cells[NomStrExcelWhile,NomPosTabled-11];
        St1:=IntToStr(NomStrExcelWhile-2)+'). '+st1;
        if FindAndReplaceRetry(StrTable,st1) then
          NomStrExcelWhile:=NomStrExcelWhile+1
        else
          NomStrExcelWhile:=FMain.SeNomBDRow.MaxValue;
        end;
      end;    }
    inc (NomPosTabled);
    end;
  Word.Documents.Close;

  Inc(NomStrPeople);
  end;
Word.Quit;
Word:=UnAssigned;
ShowMessage('Создание документов успешно выполнено');
end;
end;

procedure TFDocuments.BtAddSelectPeopleAllClick(Sender: TObject);
var
i:LongWord;
begin
i:=0;
While i<LbAllPeople.Count do
  begin
  MeSelectPeople.Lines.Add(LbAllPeople.Items[i]);
  Inc(i);
  end;
end;

end.
