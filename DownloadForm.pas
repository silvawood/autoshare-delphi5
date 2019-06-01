unit DownloadForm;

interface

uses

  StdCtrls, Gauges, registry, WinInet, ScktComp, ZLib, ExtCtrls, Nag,
  Controls, Classes, Forms, Windows, Messages, SysUtils, Graphics, Dialogs,
  IdBaseComponent, IdHTTP, IdAntiFreezeBase, IdAntiFreeze;

type
  TfrmDownload = class(TForm)
    btnCancel: TButton;
    lblMessage: TLabel;
    ClientSocket1: TClientSocket;
    Timer1: TTimer;
    procedure btnCancelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ClientSocket1Write(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocket1Read(Sender: TObject; Socket: TCustomWinSocket);
    procedure ClientSocket1Error(Sender: TObject; Socket: TCustomWinSocket;
      ErrorEvent: TErrorEvent; var ErrorCode: Integer);
//    procedure connectToServer;
    procedure registerSoftware;
    procedure downloadPrices;
    procedure updateSuccess;
	  procedure processPriceFile(buffer:PChar;dayPoint,bufferSize:Integer);
    procedure getPricesFromWeb;
    procedure Timer1Timer(Sender: TObject);
    procedure ClientSocket1Connect(Sender: TObject;
      Socket: TCustomWinSocket);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmDownload: TfrmDownload;

  function nextWebFile:Boolean;
	function ReceiveBuffer(Socket: TCustomWinSocket;buffer:PChar;count:Integer):Integer;
  function ipAddrToStr(ipAddrPtr:PChar):String;

implementation

uses Autoshare1;

type
  PWord=^Word;

var
  savedS:String;
  state:Integer;//bit 2=registration(reset)/prices(set)
//0=non-redirected,1=find alternative server,2=use alternative server,3=use website prices
//8=finished
  bytesReceived, downloadDays, z, bufferSize: Integer;
  inBuff:PChar;
  lastDate:TDateTime;
  fileName:String;
  ipAddr:String;
  portNo:Word;
  IdHttp1:TIdHTTP;

const
  defaultHost:String='www.silvawood.co.uk';
  defaultAddress:String='80.229.159.113';
  defaultPort:Integer=8080;

{$R *.DFM}

procedure TfrmDownload.registerSoftware;
var
  prompt:String;
begin
prompt:='Please enter your e-mail address, as used by PayPal';
S:='';
Repeat
	if Not(InputQuery('Registration: Enter E-mail Address',prompt,S)) then Exit;
	prompt:='This is not a valid e-mail address. Please try again.';
Until Pos('@',S)>0;
with TIdHTTP.Create do
  begin
  S:=Get('http://ccgi.silvawood.force9.co.uk/reg.php?s=0'+S);
  if length(S)=4 then
    begin
    //successful, so extract orderId
    orderId:=(Ord(S[1]) shl 24) or (Ord(S[2]) shl 16) or (Ord(S[3]) shl 8) or Ord(S[4]);
    end
  else
    //failed, so show error message
    ShowMessage(S);
  end;
end;

{procedure TfrmDownload.downloadPrices;
begin
//state:=4;//prices download
downloadDays:=0;
bytesReceived:=0;
//send last date along with orderId (LSBytes are first, include only 3 bytes)
S:=Chr(datesArray[market,numDays-1] and 255) + Chr(datesArray[market,numDays-1] shr 8)
 + Chr(orderId and 255)+Chr(orderId shr 8 and 255) + Chr(orderId shr 16 and 255);
//connectToServer;//automatically tries alternatives
end;

procedure TfrmDownload.connectToServer;
var
  buffer: PChar;
begin
InternetAttemptConnect(0);
//see if an alternative server is in use
ipAddr:=defaultAddress;
portNo:=defaultPort;
With TRegistry.Create do
  begin
  RootKey:=HKEY_CURRENT_USER;
  if OpenKey(regKey,false) and ValueExists('Key') then
    begin
    GetMem(buffer,6);
    try
      ReadBinaryData('Key', buffer^, 6);
      ipAddr:=ipAddrToStr(buffer+2);
      portNo:=PWord(buffer)^ xor 65535;
	  finally
	 	  Free;
      FreeMem(buffer,6)
      end;
    end;
  end;
with ClientSocket1 do
  begin
  Port:=portNo;
  Address:=ipAddr;
  try
    Open;
  except
    MessageDlg('Could not connect to the internet. Please check that a firewall is not blocking communications.',
    mtError, [mbOK], 0);
    Exit;
    end;
  end;
ShowModal;
end; }

procedure TfrmDownload.btnCancelClick(Sender: TObject);
begin
IdHttp1.Disconnect;
//Form1.restorePrices;
frmDownload.Close;
{if state<8 then //cancel if download not already finished
  begin
  state:=8;//stop further download
  frmDownload.Close;
  Form1.restorePrices;
  end;}
end;

procedure TfrmDownload.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
ClientSocket1.Close;
ClientSocket1.Free;
FreeMem(inBuff);
Action:=caFree;
end;

procedure TfrmDownload.ClientSocket1Write(Sender: TObject;
  Socket: TCustomWinSocket);
begin
S:=Copy(S,Socket.SendText(S)+1,65535);//reduce S to unsent chrs
end;

procedure TfrmDownload.ClientSocket1Read(Sender: TObject;
  Socket: TCustomWinSocket);
var
  a:Integer;
  newIpAddr:String;
  newPortNo:Word;
begin
Application.ProcessMessages;
if state>=8 then Exit;//finished
if (state and 3)=1 then //try to find alternative server
  begin
  S:=S+Socket.ReceiveText;
  a:=Pos(#13#10#13#10,S);
  if (a=0) or (Length(S)<a+9) then Exit;
  newIpAddr:=ipAddrToStr(PChar(S)+a+5);
  newPortNo:=(PWord(PChar(S)+a+3))^ xor 65535;
  if (portNo=newPortNo) and (ipAddr=newIpAddr) then //already using redirection
    begin
    if state>3 then//prices download, so try web site
      getPricesFromWeb
    else//registration, so give up
      begin
      frmDownload.Close;
      MessageDlg('Could not connect to server. Please check your internet connection and then try again.',
    mtError, [mbOK], 0);
      end;
    end
  else with ClientSocket1 do
    begin
    Close; //close socket
    Address:=newIpAddr;
    Port:=newPortNo;
    inc(state);//state of 2 causes connection to alternative ip
    with TRegistry.Create do //save new settings
      try
		    RootKey:=HKEY_CURRENT_USER;
		    if OpenKey(regKey,true) then
			    WriteBinaryData('Key', (PChar(S)+a+3)^, 6);
	    finally
		    Free;
      end;
    Host:='';
    S:=savedS;//re-fetch string to send to server
    Open;//try again with new connection
    end;
  end
else if (state and 4)=0 then //registration
  begin
  S:=S+Socket.ReceiveText;
  if length(S)<3 then Exit;//still more to come
  state:=8;//finished
  Close;//close form
  Case (Ord(S[1]) and 3) of
	  0:  begin
    	  orderId:=((Ord(S[1]) shr 2) or (Ord(S[2]) shl 6) or (Ord(S[3]) shl 14));
   		  ShowMessage('Registration was successful and you are now using the registered version of this software.');
        end;
	  1: MessageDlg('The e-mail address you entered was not correct. It must be the one you use with PayPal.',
    mtError, [mbOK], 0);
    2: MessageDlg('This software has already been registered several times.',
    mtError, [mbOK], 0);
	  end;
  end
else if state=7 then//get price files from website
  begin
  if bufferSize=0 then
    begin
    S:=S+Socket.ReceiveText;
    z:=Pos(#13#10#13#10,S);
    if (z=0) then Exit;//not yet received all of header
    if Pos('Length: ',S)>0 then //price file was found on web server
      begin
      bufferSize:=StrToInt(Copy(S, Pos('Length: ',S) + length('Length: '), 4))+z+3;
      lblMessage.Caption:='Downloading prices for '
      +Copy(fileName,5,2)+'/'+Copy(fileName,3,2)+'/'+Copy(fileName,1,2);
      end;
    end;
  if bufferSize>0 then
    begin
    S:=S+Socket.ReceiveText;
    if length(S)<bufferSize then Exit;
    dec(bufferSize,z+5);//also exclude two bytes of outputsize
	  Move(datesArray[1],datesArray[0],2*(numDays-1));
    datesArray[market,numDays-1]:=(StrToInt(Copy(fileName,1,2))shl 9)
      +(StrToInt(Copy(fileName,3,2))shl 5)+(StrToInt(Copy(fileName,5,2)));
	  //move up company prices to make way for new day
    a:=0;
    while a<numCompanies[market] do
	    begin
	    Move(pricesArray[a,1],pricesArray[a,0],2*(numDays-1));
      pricesArray[market,a,numDays-1]:=0;
	    inc(a);
	    end;
//    processPriceFile(PChar(S)+z+3,numDays-2);
    end;
  if nextWebFile then
    Timer1.Enabled:=true //timer event will re-open socket connection
  else
    updateSuccess;
	end
else//get price files from server
  try
  if downloadDays=0 then //not yet reading files
    begin
    if bytesReceived=0 then GetMem(inBuff,8192);
    inc(bytesReceived, ReceiveBuffer(Socket,inBuff+bytesReceived,2));
    if bytesReceived<2 then Exit;//receive rest in next Read event
    if PByte(inBuff)^>0 then //error has occurred
      begin
      state:=8;//finish download
      a:=PByte(inBuff)^;
      Hide;
      Case a of
        1: MessageDlg('Today''s closing share prices are not yet available.',
    mtError, [mbOK], 0);
  	    2: begin
           orderId:=0;
           MessageDlg('This product has not been properly registered.',
    mtError, [mbOK], 0);
           end;
  	    3: expired:=true;
        end;
      Close;
      Exit;
      end
    else //so far, so good
      begin
      downloadDays:=PByte(inBuff+1)^;
      //move up dates array to make way for new days
	    Move(datesArray[downloadDays],datesArray[0],2*(numDays-downloadDays));
	    //move up company prices to make way for new days
	    a:=0;
      while a<numCompanies[market] do
		    begin
		    Move(pricesArray[a,downloadDays],pricesArray[a,0],2*(numDays-downloadDays));
	      FillChar(pricesArray[a,numDays-downloadDays],2*downloadDays,0);//clear new days
		    inc(a);
		    end;
      bytesReceived:=0;
      bufferSize:=0;
      z:=downloadDays;
      setLength(S,8);
      end;
    end;
  if bufferSize=0 then//ready to receive header info for next file
    begin
    inc(bytesReceived,ReceiveBuffer(Socket,PChar(S)+bytesReceived,8));
    if bytesReceived<8 then Exit;
    lblMessage.Caption:='Downloading prices for day '+IntToStr(downloadDays-z+1)+' of '+IntToStr(downloadDays);
    datesArray[market,numDays-z]:=(StrToInt(Copy(S,3,2))shl 9)
    +(StrToInt(Copy(S,5,2))shl 5)+(StrToInt(Copy(S,7,2)));
    bufferSize:=Ord(S[1]) or (Ord(S[2]) shl 8)-8;
    bytesReceived:=0;
    end;
  //now receive prices file into buffer
  inc(bytesReceived,ReceiveBuffer(Socket,inBuff+bytesReceived,bufferSize-bytesReceived));
  if bytesReceived<bufferSize then Exit;
  bytesReceived:=0;
//  processPriceFile(inBuff,numDays-1-z);
  bufferSize:=0;
  dec(z);
  if z=0 then updateSuccess;
  except
    btnCancelClick(Sender);
    MessageDlg('An error occurred during the updating of prices. Please try again.',
    mtError, [mbOK], 0);
  end;
end;

procedure TfrmDownload.ClientSocket1Error(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
  var ErrorCode: Integer);
begin
ErrorCode:=0;
//Form1.restorePrices;//to prevent corruption
if ((state and 3)=0)
  and InternetCheckConnection(PChar('http://'+defaultHost), 1, 0) then
  //try to find alternative server
  begin
  lblMessage.Caption:='Seeking alternative server';
  inc(state);//state of 1 indicates fetching alternative ip
  savedS:=S;//store for use later (for registration)
  with ClientSocket1 do
    begin
    Close;
    Host:=defaultHost;
    S:='GET /blank.htm HTTP/1.0'#13#10 + 'Host: ' + defaultHost + #13#10#13#10;
    Port:=80;
    Open;//try opening connection to website
    end;
  end
else //could not connect to alternative server either
  begin
  if ((state and 4)=0){registration} or not InternetCheckConnection(PChar('http://'+defaultHost), 1, 0) then
    begin
    //already tried alternative, so give up if registration or if cannot contact web site
    Close;//close form
    MessageDlg('Could not connect to server. Please check your internet connection and then try again.',
    mtError, [mbOK], 0);
    end
  else
    getPricesFromWeb;//prices download
  end
end;

procedure TfrmDownload.processPriceFile(buffer:PChar;dayPoint,bufferSize:Integer);
var
  outputSize,i,j,k,nonZeroDays:Integer;
  P,outBuff:Pointer;
  PNames:PChar;
  newDelisted,n,newShares,shareNo: Word;
  tempArray:array[0..749] of Word;
  cName:String;
begin
//dayPoint used for corrections, and is actually current day -1
outputSize:=PWord(buffer)^;
DecompressBuf(Pointer(buffer+2),bufferSize,outputSize,outBuff,j);
P:=outBuff;
//first,process any corrections to previous day's prices
//get number of corrections
j:=PWord(P)^;
inc(PWord(P));
//process corrections
While j>0 do
	begin
  //get share no and price
  shareNo:=PWord(P)^;
  inc(PWord(P));
  //store corrected price, remembering to account for share number
  //discrepancy between client and server, and updating 'downloadDays-i' which
  //is the location of the previous day's price
  pricesArray[market,shareNo+numDelisted[market],dayPoint]:=PWord(P)^;
  inc(PWord(P));
  dec(j);
  end;
//now read number of delisted shares
newDelisted:=PWord(P)^;
inc(PWord(P));
j:=0;
PNames:=PChar(namesArray[market]);
While j<newDelisted do
	begin
  shareNo:=PWord(P)^;
  inc(PWord(P));
  //delisted shares are not included in share numbering on server side
  inc(shareNo,numDelisted[market]);
  cName:=companyName(market,shareNo);
  Form1.lstDelisted.Items.Add(cName);
  //swap share with that immediately following end of delisted list
  //at beginning of prices array/shareNames string
  Move((PNames+(numDelisted[market]+j)*nameLength)^,(PNames+shareNo*nameLength)^,nameLength);
  Move(PChar(cName)^,(PNames+(numDelisted[market]+j)*nameLength)^,nameLength);
  Move(pricesArray[shareNo,0],tempArray,numDays*2);
  Move(pricesArray[numDelisted[market]+j,0],pricesArray[market,shareNo,0],2*numDays);
  Move(tempArray,pricesArray[numDelisted[market]+j,0],2*numDays);
  inc(j);
  end;
//read number of new shares
newShares:=PWord(P)^;
inc(PWord(P));
inc(numCompanies[market],newShares);
setLength(pricesArray,market,numCompanies[market]);
//add new share names
j:=Length(namesArray[market]);
setLength(namesArray[market],j+nameLength*newShares);
Move(P^,(PChar(namesArray[market])+j)^,nameLength*newShares);
//skip over string portion of buffer
inc(PChar(P),nameLength*newShares);
//combine any delisted shares and new issues having the same name
//otherwise clear new issue elements so that previous day's prices are zero
j:=numCompanies[market]-newShares;
while j<numCompanies[market] do
	begin
  k:=0;
  while (k<numDelisted[market]) and (companyName(market,j)<>companyName(market,k)) do inc(k);
  if k=numDelisted[market] then
    begin
  	FillChar(pricesArray[market,j,0],2*numDays,0); //share doesn't already exist
    inc(j);
    end
  else
  	begin
    n:=numDays-1;
    while pricesArray[market,k,n]=0 do dec(n);
    //fill gap in prices
    for i:=n to numDays-1 do pricesArray[market,j,i]:=pricesArray[market,k,n];
    //move delisted (k) price to new company (j) price for each day
  	for n:=n-1 downto 0 do pricesArray[market,j,n]:=pricesArray[market,k,n];
    //delete delisted share (can't have two shares with the same name)
    cName:=companyName(market,k);
//    Form1.removeFromPortfolio(cName);
    Form1.lstDelisted.Items.Delete(Form1.lstDelisted.Items.IndexOf(cName));
    deleteShare(market,k);//also decs numCompanies
    dec(numDelisted[market]);
    //j is effectively incremented when delisted share is deleted and shares moved down
    end;
	end;
//read day's prices for all shares (new and existing)
j:=numDelisted[market]+newDelisted; //skip delisted shares, which are not transmitted
shareNo:=j;
inc(dayPoint);//advance to current day, to save having dayPoint+1 recalculated in loop
While j<numCompanies[market] do
	begin
  pricesArray[market,j,dayPoint]:=PWord(P)^;
  inc(PWord(P));
  inc(j);
  end;
//auto-delete any new delisted shares that have only one day of price data
While shareNo>numDelisted[market] do
	begin
  dec(shareNo);
	nonZeroDays:=0;
  j:=numDays-1;
	Repeat
		if pricesArray[market,shareNo,j]>0 then
    	begin
  		inc(nonZeroDays);
    	if nonZeroDays>1 then Break;
    	end;
    dec(j);
	Until j<0;
  if nonZeroDays<=1 then
  	begin
    cName:=companyName(market,shareNo);
    Form1.removeFromPortfolio(cName);
    Form1.lstDelisted.Items.Delete(Form1.lstDelisted.Items.IndexOf(cName));
    deleteShare(market,shareNo);
    dec(newDelisted);
    end;
	end;
//now numDelisted can be updated
inc(numDelisted[market],newDelisted);
FreeMem(outBuff);
end;

procedure TfrmDownload.getPricesFromWeb;
var
  year,month,day:Word;
begin
state:=7;//get prices from web site
//use socket connection to website to retrieve price files
DecodeDate(Date,year,month,day);
lastDate:=EncodeDate(year,datesArray[market,numDays-1] shr 5 and 15,datesArray[market,numDays-1] and 31);
if lastDate>Date then lastDate:=incMonth(lastDate,-12);
if nextWebFile then with ClientSocket1 do
  begin
  Close;
  Address:='';
  Host:=defaultHost;
  Port:=80;
  Timer1.Enabled:=true;
//  Open;//start connection
  end
else
  begin
  Close;//close form
  ShowMessage('Your prices are already up-to-date.');
  end;
end;

function nextWebFile:Boolean;
var
  year,month,day:Word;
begin
Repeat
  lastDate:=lastDate+1;
Until ((DayOfWeek(lastDate)<>1) and (DayOfWeek(lastDate)<>7)) or (lastDate>Date);
Result:=lastDate<=Date;
if Result then
  begin
  DecodeDate(lastDate,year,month,day);
  fileName:=twoDigits(year-2000)+twoDigits(month)+twoDigits(day);
  S:='GET /Prices/' + fileName + '.dat' + ' HTTP/1.0'#13#10 + 'Host: ' + defaultHost + #13#10#13#10;
  bufferSize:=0;
  end
end;

procedure TfrmDownload.updateSuccess;
begin
state:=8;//finished download
lblMessage.Caption:='Updating lists';
Application.ProcessMessages;
With Form1 do
	begin
  pricesChanged:=true;
	btnSave.Enabled:=True;
	btnUndo.Enabled:=True;
	btnSelectAll.Enabled:=(numDelisted>0);
  refreshLists;
	end;
Close;//close form
end;

function ipAddrToStr(ipAddrPtr:PChar):String;
type
	PCardinal=^Cardinal;
var
  ipAddr:Cardinal;
begin
ipAddr:=PCardinal(ipAddrPtr)^ xor 4294967295;
Result:=IntToStr(ipAddr and 255)+'.'+IntToStr(ipAddr shr 8 and 255)+'.'
+IntToStr(ipAddr shr 16 and 255)+'.'+IntToStr(ipAddr shr 24 and 255)
end;

function ReceiveBuffer(Socket: TCustomWinSocket;buffer:PChar;count:Integer):Integer;
begin
Result:=Socket.ReceiveBuf(buffer^,count);
if Result=-1 then Result:=0;
end;

procedure TfrmDownload.Timer1Timer(Sender: TObject);
begin
if not ClientSocket1.Active then
  begin
  ClientSocket1.Open;
  Timer1.Enabled:=false;
  end;
end;

procedure TfrmDownload.ClientSocket1Connect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
if ((state and 3)=0) or ((state and 3)=2) then lblMessage.Caption:='Talking with server';
end;

end.
