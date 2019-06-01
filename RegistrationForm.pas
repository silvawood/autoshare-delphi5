unit RegistrationForm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, WinInet, regthread, DownloadThread, ComCtrls, Registry, IdHTTP;

type
  TfrmRegistration = class(TForm)
    btnCancel: TButton;
    Memo1: TMemo;
    ProgressBar1: TProgressBar;
    procedure btnCancelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure registerSoftware(softwareID:Integer);
    procedure downloadPrices;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRegistration: TfrmRegistration;
  downloadThread1: TGetPrices;
  regThread1: HTTPThread;
  procedure setProxyOptions(httpConn:TIdHTTP);
  procedure toggleControls(state:Boolean);

implementation

uses AutoShare1, Nag;

{$R *.DFM}

procedure TfrmRegistration.btnCancelClick(Sender: TObject);
begin
if btnCancel.Caption='&Cancel' then
  begin
  Caption:='Cancelling';
  btnCancel.Enabled:=false;
  Application.ProcessMessages;
  http1.Disconnect;
  if downloadThread1<>nil then
    begin
    downloadThread1.Terminate;
    downloadThread1.WaitFor;//but freezes if Cancel clicked before
    end
  else if regThread1<>nil then
    begin
    regThread1.Terminate;
    regThread1.WaitFor;
    end;
  end;
Close;
end;

procedure TfrmRegistration.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
toggleControls(true);
Form1.btnSave.Enabled:=portfolioChanged or anyPricesChanged;
http1.Free;
Action:=caFree;
if expired then TfrmNag.Create(self).ShowModal;
end;

procedure TfrmRegistration.registerSoftware(softwareID:Integer);
var
  prompt, regStr:String;
begin
prompt:='Please enter your PayPal e-mail address:';
regStr:='';
Repeat
	if Not(InputQuery('Upgrading',prompt,regStr)) then Exit;
	prompt:='This is not a valid e-mail address. Please try again.';
Until Pos('@',regStr)>0;
InternetAttemptConnect(0);
Show;
regThread1:=HTTPThread.Create(IntToStr(softwareID)+LowerCase(regStr));
end;

procedure TfrmRegistration.downloadPrices;
begin
Caption:='Updating prices';
toggleControls(false);
Form1.btnSave.Enabled:=false;
InternetAttemptConnect(0);
Show;
downloadThread1:=TGetPrices.Create(false);
end;

procedure toggleControls(state:Boolean);
begin
with Form1 do
  begin
  lstMarkets.Enabled:=state;
  btnCombine.Enabled:=state;
  btnDelete.Enabled:=state;
  btnUndo.Enabled:=state;
  btnUpdate.Enabled:=state;
  btnImport.Enabled:=state;
  btnExit.Enabled:=state;
  end;
end;

procedure setProxyOptions(httpConn:TIdHTTP);
var
  options:Integer;
begin
With TRegistry.Create do
  begin
  try
	RootKey:=HKEY_CURRENT_USER;
	if OpenKey('Software\Silvawood\Proxy',false) and (KeyExists('Options')) then with http1.ProxyParams do
    begin
    options:=ReadInteger('Options');
    if (options and 1) > 0 then
      begin
      //use proxy settings in Silvawood registry entry
		  httpConn.ProxyParams.ProxyServer:=ReadString('Address');
      httpConn.ProxyParams.ProxyPort:=options shr 2;
      if (options and 2) > 0 then
        begin
        //authorization required
        httpConn.ProxyParams.ProxyUsername:=ReadString('Username');
        httpConn.ProxyParams.ProxyPassword:=ReadString('Password');
        end;
      end;
    end;
	except end;
  Free;
  end;
end;

end.
