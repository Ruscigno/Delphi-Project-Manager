program DelphiProjectManager;

uses
  Forms,
  udpmMainForm in 'udpmMainForm.pas' {fdpmMainForm};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Delphi Project Manager';
  Application.CreateForm(TfdpmMainForm, fdpmMainForm);
  Application.Run;
end.
