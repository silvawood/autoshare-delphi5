unit Proxy;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Registry, Autoshare1, HH, HH_FUNCS;

type
  TfrmProxy = class(TForm)
    lblAddress: TLabel;
    edtAddress: TEdit;
    lblPort: TLabel;
    edtPort: TEdit;
    chkAuthorization: TCheckBox;
    btnOK: TButton;
    btnCancel: TButton;
    grpAuthorization: TGroupBox;
    lblUsername: TLabel;
    edtUsername: TEdit;
    lblPassword: TLabel;
    edtPassword: TEdit;
    chkUseProxy: TCheckBox;
    btnHelp: TButton;
    procedure chkAuthorizationClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure chkUseProxyClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure btnHelpClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmProxy: TfrmProxy;

implementation

{$R *.DFM}

procedure TfrmProxy.chkAuthorizationClick(Sender: TObject);
begin
if chkAuthorization.Checked then
  begin
  grpAuthorization.Show;
  Height:=279;
  btnOK.Top:=216;
  btnCancel.Top:=216;
  btnHelp.Top:=216;
  end
else
  begin
  grpAuthorization.Hide;
  Height:=179;
  btnOK.Top:=116;
  btnCancel.Top:=116;
  btnHelp.Top:=116;
  end
end;

procedure TfrmProxy.btnOKClick(Sender: TObject);
var
  portNo, error:Integer;
begin
Val(edtPort.Text,portNo,error);
if chkUseProxy.Checked then
  begin
  if (error>0) or (portNo<0) or (portNo>65535) then
    begin
    MessageDlg('Invalid port number.', mtError, [mbOK], 0);
    Exit;
    end;
  if chkAuthorization.Checked then
    begin
    if edtUsername.Text='' then
      begin
      MessageDlg('Username missing.', mtError, [mbOK], 0);
      Exit;
      end;
    if edtPassword.Text='' then
      begin
      MessageDlg('Password missing.', mtError, [mbOK], 0);
      Exit;
      end;
    end;
  end;
With TRegistry.Create do
  begin
	try
		RootKey:=HKEY_CURRENT_USER;
		if OpenKey('Software\Silvawood\Proxy',true) then
			begin
  	  WriteString('Address',edtAddress.Text);
      WriteInteger('Options',Ord(chkUseProxy.Checked)  or (Ord(chkAuthorization.Checked) shl 1)
         or (portNo shl 2));
      WriteString('Username',edtUsername.Text);
      WriteString('Password',edtPassword.Text);
      end
	except end;
  Free;
  end;
Close;
end;

procedure TfrmProxy.btnCancelClick(Sender: TObject);
begin
Close;
end;

procedure TfrmProxy.chkUseProxyClick(Sender: TObject);
var
  useProxy:Boolean;
begin
useProxy:=chkUseProxy.Checked;
if not useProxy then chkAuthorization.Checked:=false;
chkAuthorization.Enabled:=useProxy;
lblAddress.Enabled:=useProxy;
edtAddress.Enabled:=useProxy;
lblPort.Enabled:=useProxy;
edtPort.Enabled:=useProxy;
end;

procedure TfrmProxy.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action:=caFree;
end;

procedure TfrmProxy.FormCreate(Sender: TObject);
var
  options:Integer;
begin
With TRegistry.Create do
  begin
	try
		RootKey:=HKEY_CURRENT_USER;
		if OpenKey('Software\Silvawood\Proxy',true) and ValueExists('Options') then
			begin
      options:=ReadInteger('Options');
      chkUseProxy.Checked:=(options and 1) > 0;
      chkAuthorization.Checked:=(options and 2)>0;
      edtPort.Text:=IntToStr(options shr 2);
			edtAddress.Text:=ReadString('Address');
      edtUsername.Text:=ReadString('Username');
      edtPassword.Text:=ReadString('Password');
      end;
	except end;
  Free;
  end;
chkUseProxyClick(Sender);
chkAuthorizationClick(Sender);
end;

procedure TfrmProxy.btnHelpClick(Sender: TObject);
begin
helpWindow:=HtmlHelp(GetDesktopWindow, '..\AutoShare.chm', HH_HELP_CONTEXT, 800);
end;

end.
