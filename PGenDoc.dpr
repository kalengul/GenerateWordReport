program PGenDoc;

uses
  Forms,
  GeneratorDoc in 'GeneratorDoc.pas' {FMain},
  padegFIO in 'padegFIO.pas',
  PageDoc in 'PageDoc.pas' {FDocuments};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFMain, FMain);
  Application.CreateForm(TFDocuments, FDocuments);
  Application.Run;
end.
