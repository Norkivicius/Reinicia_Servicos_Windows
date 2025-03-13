program Agendar_Programa;

uses
  Vcl.Forms,
  Unit1 in 'Unit1.pas' {Frm_Main},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFrm_Main, Frm_Main);
  Application.Run;

end.
