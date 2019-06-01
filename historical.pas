unit historical;

interface

uses
  Classes, IdBaseComponent, IdHTTP, SysUtils, Dialogs, ZLib, IdComponent,
  IdTCPConnection, IdTCPClient, hyperStr;

type
  TGetHistorical = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
    procedure ThreadTerminated(Sender:TObject);
  end;

implementation

uses Autoshare1, DownloadThread, RegistrationForm,Windows;

procedure TGetHistorical.Execute;
var
  S,localFromDateUrl,localToDateUrl:String;
  i,j,k,dateNo,priceDateNo,companyLen,start,oldPriceDateNo:Integer;
  httpLocal:TIdHTTP;
  localShareNo:Integer;
  year,month,day,dateWord,lastDate:Word;
  adjClose:Boolean;
  price:Single;
begin
OnTerminate:=ThreadTerminated;
FreeOnTerminate:=true;
//send last date along with orderId (MSB first, include only 3 bytes)
httpLocal:=TIdHTTP.Create;
httpLocal.HandleRedirects:=true;
setProxyOptions(httpLocal);
adjClose:=(commaCount<>3);
Repeat
  EnterCriticalSection(CS1);
  inc(globalShareNo);
  localShareNo:=globalShareNo;
  frmRegistration.ProgressBar1.Position:=globalShareNo;
//  frmRegistration.Memo1.Text:=inttostr(globalshareno);
  LeaveCriticalSection(CS1);
  if localShareNo>=noCompanies then
    begin
    Terminate;
    Break;
    end;
  companyLen:=length(yahooPrices[market,localShareNo]);
//  dateOffset:=companyLen-length(datesArray[market]);
  if (localShareNo>=numDelisted[market])
   or ((companyLen>1) and (yahooPrices[market,localShareNo,companyLen-1]>0) //delisted share to be reinstated
    and (yahooPrices[market,localShareNo,companyLen-2]<=0)) then
  begin
  if {((companyLen=0) and not latestAvailable) or (}(companyLen=1){ and latestAvailable)}
   or (yahooPrices[market,localShareNo,companyLen-(length(datesArray[market])-(oldNoDays-1))]>0) then
    begin
    localFromDateUrl:=pricesFromDateUrl;
    if companyLen=1 then
      oldPriceDateNo:=0
    else
      oldPriceDateNo:=companyLen-(length(datesArray[market])-oldNoDays);
    dateNo:=oldNoDays;
    end
  else
    begin
    oldPriceDateNo:=companyLen-2;
    while yahooPrices[market,localShareNo,oldPriceDateNo]<=0 do dec(oldPriceDateNo);
    inc(oldPriceDateNo);
    dateNo:=length(datesArray[market])-(companyLen-oldPriceDateNo);
    dateWord:=datesArray[market,dateNo];
    day:=dateWord and 31;
    month:=(dateWord shr 5) and 15;
    year:=(dateWord shr 9) + 1900;
    localFromDateUrl:='&a='+IntToStr(month-1)+'&b='+IntToStr(day)+'&c='+IntToStr(year);
    end;
  if latestAvailable and (yahooPrices[market,localShareNo,companyLen-1]=0) then
    localToDateUrl:=''
  else
    localToDateUrl:=pricesToDateUrl;
  try
    S:=httpLocal.Get('http://'+{histUrl+}mainPricesPrefix
    +TrimRight(Copy(namesArray[market],localShareNo*nameLength + 1,nameLength))
    +localToDateUrl+localFromDateUrl);
  except
    S:='';
    end;
  i:=length(S)-1;
  if (i>0) and (S[i+1]=chr(10)) then
    begin
    lastDate:=datesArray[market,length(datesArray[market])-1];
    if {((companyLen=0) and not latestAvailable) or (}(companyLen=1){ and latestAvailable)} then
      begin
      //need to uniquely size array for company
      Repeat dec(i) Until S[i]=Chr(10);
      try
        year:=StrToInt(Copy(S,i+1,4));
        month:=StrToInt(Copy(S,i+6,2));
        day:=StrToInt(Copy(S,i+9,2));
        dateWord:=dateToWord(day,month,year);
      except dateWord:=0; end;
      dateNo:=length(datesArray[market])-1;
      if datesArray[market,0]<=dateWord then
        begin
        while datesArray[market,dateNo]>dateWord do dec(dateNo);
        if datesArray[market,dateNo]<>dateWord then inc(dateNo);
        SetLength(yahooPrices[market,localShareNo],length(datesArray[market])-dateNo);
        end
      else
        SetLength(yahooPrices[market,localShareNo],dateNo+1);
      if latestAvailable then
        yahooPrices[market,localShareNo,length(yahooPrices[market,localShareNo])-1]:=yahooPrices[market,localShareNo,0];
      dateNo:=0;
      while (datesArray[market,dateNo]<dateWord) do inc(dateNo);
      i:=length(S)-1;
      end;
    priceDateNo:=oldPriceDateNo;
    start:=1;
    while S[start]<>chr(10) do inc(start);
    //if latestAvailable then dec(companyLen);
    while i>start do
      begin
      //move to start of price
      if not adjClose then
        begin
        while (S[i]<>',') do dec(i);
        dec(i);
        while (S[i]<>',') do dec(i);
        dec(i);
        end;
      j:=i;
      Repeat dec(i) Until S[i]=',';
      price:=strToPrice(Copy(S,i+1,j-i));
      //if price>10000 then price:=price/100;
      while S[i]<>chr(10) do dec(i);
      try
        year:=StrToInt(Copy(S,i+1,4));
        month:=StrToInt(Copy(S,i+6,2));
        day:=StrToInt(Copy(S,i+9,2));
        dateWord:=dateToWord(day,month,year);
      except dateWord:=0; end;
      dec(i);
      if lastDate>=dateWord then
        begin
        while (datesArray[market,dateNo]<dateWord) {and (endDateNo>=0)} do
          begin
          if priceDateNo>0 then
            yahooPrices[market,localShareNo,priceDateNo]:=yahooPrices[market,localShareNo,priceDateNo-1];
          inc(priceDateNo);
          inc(dateNo);
          end;
        if datesArray[market,dateNo]=dateWord then
          begin
          yahooPrices[market,localShareNo,priceDateNo]:=price;
          inc(priceDateNo);
          inc(dateNo);
          end
        else //no exact match
          begin
          //search in extraDates to see if already noted
          EnterCriticalSection(CS2);
          k:=length(extraDates)-1;
          while (k>=0) and (extraDates[k]<>dateWord) do dec(k);
          if k<0 then //date not already noted
            begin
            k:=length(extraDates);
            SetLength(extraDates,k+1);
            SetLength(extraPrices,k+1,noCompanies);
            FillChar(extraPrices[k,0],4*noCompanies,0);
            extraDates[k]:=dateWord;
            end;
          extraPrices[k,localShareNo]:=price;
          LeaveCriticalSection(CS2);
          end;
        end;
      end;
    for i:=oldPriceDateNo to companyLen-2 do
      if ((yahooPrices[market,localShareNo,i-1]>2*yahooPrices[market,localShareNo,i])
      and (yahooPrices[market,localShareNo,i+1]>2*yahooPrices[market,localShareNo,i]))
      or ((yahooPrices[market,localShareNo,i]>2*yahooPrices[market,localShareNo,i-1])
      and (yahooPrices[market,localShareNo,i]>2*yahooPrices[market,localShareNo,i+1])) then
        yahooPrices[market,localShareNo,i]:=-yahooPrices[market,localShareNo,i-1];
    for i:=oldPriceDateNo to companyLen-1 do
      if yahooPrices[market,localShareNo,i]=0 then
        yahooPrices[market,localShareNo,i]:=-Abs(yahooPrices[market,localShareNo,i-1]);
    end;
  end;

  //move days down if original size was too large, as extra dates found
{  if priceDateNo>oldLen then
    begin
    inc(endDateNo);
    Move(yahooPrices[market,localShareNo,endDateNo]
    ,yahooPrices[market,localShareNo,oldLen]
    ,4*(length(yahooPrices[market,localShareNo])-endDateNo+1));
    SetLength(yahooPrices[market,localShareNo],length(yahooPrices[market,localShareNo])-(endDateNo-oldLen));
    end;}
Until Terminated;
httpLocal.Free;
end;

procedure TGetHistorical.ThreadTerminated(Sender:TObject);
begin
dec(threadsRunning);
end;


end.
 