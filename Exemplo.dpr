program Exemplo;

uses
  Forms,
  Unt_Principal in 'Unt_Principal.pas' {Frm_Principal},
  Oleacc in 'Oleacc.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'Exemplo Artigo';
  Application.CreateForm(TFrm_Principal, Frm_Principal);
  Application.Run;
end.
