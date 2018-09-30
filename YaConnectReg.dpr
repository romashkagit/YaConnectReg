program YaConnectReg;





uses
  Forms,
  Main in 'Main.pas' {YaConnectReg},
  SelDelBox in 'SelDelBox.pas' {Form2},
  Thread in 'Thread.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'YaConnectReg';
  Application.CreateForm(TYaConnectReg, YaConnect);
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.
