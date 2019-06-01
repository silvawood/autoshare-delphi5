program AutoShare;

uses
  Forms,
  AutoShare1 in 'AutoShare1.pas' {Form1},
  Nag in 'Nag.pas' {frmNag},
  DownloadThread in 'DownloadThread.pas',
  RegistrationForm in 'RegistrationForm.pas' {frmRegistration},
  regthread in 'regthread.pas',
  Proxy in 'Proxy.pas' {frmProxy},
  historical in 'historical.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.HelpFile := '';
  Application.Title := 'AutoShare';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
