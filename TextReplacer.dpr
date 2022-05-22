program TextReplacer;

uses
  Forms,
  Main in 'Main.pas' {MainForm},
  ExcelImportExport in 'ExcelImportExport.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
