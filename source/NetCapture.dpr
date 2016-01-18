program NetCapture;

uses
  Forms,
  Windows,
  UMain in 'UMain.pas' {MainForm},
  Unit2 in 'Unit2.pas' {Form2},
  Unit3 in 'Unit3.pas' {Form3};

{$R *.res}

begin
  ShowWindow(Application.Handle, SW_Hide);
  Application.ShowMainForm:=False;
  Application.Initialize;
  Application.Title := 'NetCapture - BENBAC SOFT';
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TForm3, Form3);
  Application.Run;
end.
