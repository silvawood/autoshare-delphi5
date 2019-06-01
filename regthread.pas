unit regthread;

interface

uses
  Windows,Classes, IdHTTP, Dialogs, SysUtils, IdExceptionCore;

type
  HTTPThread = class(TThread)
  procedure DisplayMessage;

  private
    { Private declarations }
  protected
    procedure Execute; override;
  public
    constructor Create(const parameter:String);
  end;

var
  S: String;
  http1: TIdHTTP;

implementation

uses AutoShare1, RegistrationForm;

constructor HTTPThread.Create(const parameter:String);
begin
inherited Create(false);
S:=parameter;
end;

procedure HTTPThread.Execute;
begin
FreeOnTerminate:=true;
http1:=TIdHTTP.Create;
setProxyOptions(http1);
try
  S:=http1.Get(mainUrl+'/regi.php?s='+S);
except
  on EIdClosedSocket do Exit;//cancelled?
  //on EAccessViolation do Exit;
  else
    try
      mainUrl:='http://'+http1.Get(urlSource);
      S:=http1.Get(mainUrl+'/reg.php?s='+S);
    except
      on EIdClosedSocket do Exit;//cancelled?
      //on EAccessViolation do Exit;
      else S:=chr(253);
    end;
end;
Synchronize(DisplayMessage);
http1.Free;
end;

procedure HTTPThread.DisplayMessage;
var
  i:Integer;
  mktName:String;
begin
with frmRegistration do
  begin
  btnCancel.Caption:='&Close';
  if length(S)=1 then
    begin
    Case Ord(S[1]) of
      255: Memo1.Text:='The e-mail address you entered could not be found. If you are sure it is the one you use with PayPal, then please wait a few minutes and try again.';
      254: Memo1.Text:='This software has already been registered several times.';
      253: Memo1.Text:='Could not connect to registration server. Please try again after checking your internet connection and firewall.';
      end;
    end
  else if length(S)>=4 then
    begin
    Memo1.Text:='Registration was successful and you are now using the registered version of this software. Please tick the indices you'
    +' require on the Indices tab and then click the Update button.';
    //untick all indices and get latest indices file
    with Form1 do
      begin
      for i:=0 to lstMarkets.Items.Count-1 do
        if lstMarkets.Checked[i] then
          begin
          lstMarkets.Checked[i]:=false;
          mktName:=TrimRight(Copy(marketSymbols,i*nameLength + 1,nameLength));
          FileSetAttr(mktName+'.mkt',faReadOnly);//flag the file as read-only
          end;
      numMarkets:=0;
      allocMarketSpace(0);
      refreshLists;
      orderId:=(Ord(S[1]) shl 24) or (Ord(S[2]) shl 16) or (Ord(S[3]) shl 8) or Ord(S[4]);
      DeleteFile(marketsFileName);
      //save compressed market list to file
      TFileStream.Create(marketsFileName,fmCreate).Free;
      With TFileStream.Create(marketsFileName,fmOpenWrite or fmShareDenyNone) do
        try
          Write(S[5],length(S)-4);
        finally
          Free;
        end;
      processMarketsFile;
      end;
    end;
  end;
end;

end.
