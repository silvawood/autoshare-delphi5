unit Nag;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ShellAPI;

type
  TfrmNag = class(TForm)
    btnPaypal: TButton;
    Memo1: TMemo;
    btnCancel: TButton;
    lstCurrency: TComboBox;
    lblCurrency: TLabel;
    procedure btnPaypalClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmNag: TfrmNag;

  regPrice: String;

implementation

uses Autoshare1;

{$R *.DFM}

procedure TfrmNag.btnPaypalClick(Sender: TObject);
var
  url:String;
begin
url:='https://www.paypal.com/cgi-bin/webscr'
//url:='https://www.sandbox.paypal.com/cgi-bin/webscr'
+ '?cmd=_xclick'
+ '&business=EGVUHNMA5AGKL'
//+ '&business=EFZT5ZEL6VX5W'
+ '&item_name=AutoShare'
+ '&item_number=' + inttostr(orderId)
+ '&amount=' + Copy(lstCurrency.Text,2,5)
+ '&no_shipping=1'
+ '&cancel_return=http://www.silvawood.co.uk/cancelled.htm'
+ '&no_note=1'
+ '&currency_code=' + Copy(lstCurrency.Text,9,3);
ShellExecute(0,'open',PChar(url),nil,nil,SW_NORMAL);
expired:=false;
end;

procedure TfrmNag.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Action:=caFree;
end;

procedure TfrmNag.FormCreate(Sender: TObject);
begin
if expired then
  begin
  Caption:='Renewing your annual licence';
  btnPayPal.Caption:='Renew Now';
  Memo1.Text:='Your annual licence for Autoshare has expired.'#13#10#13#10
  +'To renew your licence, please select the currency in which you want to pay and then click the Renew Now button. '
  +'As an appreciation of your continuing custom, we''ve taken 10% off the usual price.'#13#10#13#10
  +'A few minutes after renewing, you will receive a confirmation e-mail and be able to '
  +'update prices again.';
  with lstCurrency.Items do
    begin
    Add('$71.95 (USD)');
    Add('£35.95 (GBP)');
    Add('€44.95 (EUR)');
    end
  end
else
  begin
  Memo1.Text:='This unregistered version of AutoShare cannot update share prices or perform backtests. '
+'If you like the software, then why not upgrade? Then you will be able to download the current and historical prices '
+'of shares from around the world, keep your prices up-to-date daily at the click of a button, and perform '
+'backtests to determine the optimum settings for maximum profits.'#13#10#13#10
+'The full version costs just $79.95 (USD), £39.95 (GBP), or €49.95 (Euro) for a year''s licence. Purchase is free '
+'from risk: if at any time during the year you decide that '
+'AutoShare is not for you, you can get a refund for your unused time.'#13#10#13#10
+'To upgrade, simply use the currency box below to select the currency in which you want to pay '
+'and then click the Upgrade Now button. You will '
+'be taken to PayPal, a well-renowned, secure and reliable '
+'payment system. You do not need to join PayPal in order to pay by credit or debit card. '
+'After you have paid, click the Close button and follow the e-mailed instructions '
+'to complete the upgrade.'#13#10#13#10
+'If you want to continue with evaluation, just click the Close button.';
  with lstCurrency.Items do
    begin
    Add('$79.95 (USD)');
    Add('£39.95 (GBP)');
    Add('€49.95 (EUR)');
    end
  end;
if CurrencyString='£' then
  lstCurrency.ItemIndex:=1
else if CurrencyString='€' then
  lstCurrency.ItemIndex:=2
else
  lstCurrency.ItemIndex:=0;
end;

procedure TfrmNag.FormShow(Sender: TObject);
var
  Rect1, Rect2: TRect;
  memoHeight: Integer;
  S: String;
begin
S:=Memo1.Text;
Memo1.Perform(EM_GETRECT, 0, longint(@rect1));
rect2 := rect1;
canvas.font := Memo1.font;
DrawTextEx(canvas.handle, Pchar(S), Length(S), rect2,
   DT_CALCRECT or DT_EDITCONTROL or DT_WORDBREAK or
   DT_NOPREFIX, Nil);
memoHeight:=Memo1.Height + rect2.Bottom - rect1.bottom;
Memo1.Height := memoHeight;
btnPayPal.Top:=memoHeight+15;
btnCancel.Top:=memoHeight+52;
lstCurrency.Top:=memoHeight+17;
lblCurrency.Top:=memoHeight+19;
Height:=memoHeight+113;
end;

end.
