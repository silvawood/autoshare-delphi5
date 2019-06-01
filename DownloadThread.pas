unit DownloadThread;

interface

uses
  Windows, Classes, IdBaseComponent, IdHTTP, SysUtils, Dialogs, Controls, ZLib,
   IdComponent, IdTCPConnection, IdTCPClient, Forms, IdExceptionCore, hyperStr;

type
  TGetPrices = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
    procedure displayMessage;
    procedure initProgressBar;
    procedure initializeHTTP;
    procedure http1Work(Sender: TObject; AWorkMode: TWorkMode; AWorkCount: Integer);
    procedure updateProgressBar;
    procedure changeButton;
    procedure restorePrices;
    procedure downloadMarketPrices;
    procedure closeDialog;
    procedure wrapUpMarkets;
    procedure downloadSharePrices;
  end;

var
  downloadSize,threadsRunning,globalShareNo, market,noCompanies,commaCount,{endDateNo,}oldNoDays: Integer;
  http1: TIdHTTP;
  msg,problems,pricesToDateUrl,pricesFromDateUrl,latestUrl,histUrl: String;
  CS1,CS2:TRTLCriticalSection;
  yearNow,monthNow,dayNow:Word;
  indexOnly,latestAvailable:Boolean;
  extraDates: array of Word;
  extraPrices: array of array of Single;

  procedure readDate(const yahooStr:String; var i:Integer; out year,month,day:Word);

implementation

uses Autoshare1, RegistrationForm,historical;

procedure TGetPrices.Execute;
var
  downloadDays: Integer;
  year,month,day:Word;
  url,serverStr:String;
begin
FreeOnTerminate:=true;
//send last date along with orderId (MSB first, include only 3 bytes)
msg:='Communicating with server';
initializeHTTP;
url:='/getPrices431.php?t='+IntToStr(orderId)+'&m=';
if FileExists(marketsFileName) then
  begin
  DecodeDate(FileDateToDateTime(FileAge(marketsFileName)),year,month,day);
  url:=url+twoDigits(year-2000)+twoDigits(month)+twoDigits(day);
  end
else
  url:=url+'000000';
try
  serverStr:=http1.Get(mainUrl+url);
except
  on EIdClosedSocket do Exit;//cancelled?
//  on EAccessViolation do Exit;//no longer necessary?
  else
    try
      mainUrl:='http://'+http1.Get(urlSource);
      serverStr:=http1.Get(mainUrl+url);
    except
      on EIdClosedSocket do Exit;//cancelled?
//      on EAccessViolation do Exit;
      else serverStr:=chr(253);
    end;
  end;
problems:='';
if length(serverStr)>0 then
  downloadDays:=Ord(serverStr[1])
else
  downloadDays:=0;
if downloadDays>=253 then
  begin
  Synchronize(changeButton);
  Case downloadDays of
    255:  begin
          msg:='This product has not been properly registered.';
          orderId:=0;
          end;
    254:  begin
          msg:='Your registration period has expired.';
          expired:=true;
          //orderId:=0;
          end;
    253:  msg:='Could not connect to server. Please try again after checking your internet connection and firewall.';
    end;
  Synchronize(displayMessage);
  end
else
  begin
  market:=0;
  if Form1.chkUnadjusted.Checked then commaCount:=3 else commaCount:=5;
  indexOnly:=Form1.chkIndexOnly.Checked;
  while market<numMarkets do
    begin
    if marketNotUTD[market] then
      begin
      if markets[market]<length(urlCountry) then //don't update EODData markets
        downloadMarketPrices
      else
        downloadSharePrices;//prices for individual imported share
      if Terminated then
        begin
        Synchronize(restorePrices);
        Exit;
        end;
      end;
    inc(market);
    end;
  Synchronize(closeDialog);
  if length(serverStr)>1 then //markets file follows
    try
    //save compressed market list to file
      TFileStream.Create(newMarketsFileName,fmCreate).Free;
      With TFileStream.Create(newMarketsFileName,fmOpenWrite or fmShareDenyNone) do
        try
          Write(serverStr[1],length(serverStr));
        finally
          Free;
        end;
      Synchronize(wrapUpMarkets);
    except
      DeleteFile(newMarketsFileName);
    end;
  end;
end;

procedure TGetPrices.initializeHTTP;
begin
http1:=TIdHTTP.Create;
setProxyOptions(http1);
http1.OnWork:=http1Work;
Synchronize(displayMessage);
end;

procedure TGetPrices.http1Work(Sender: TObject;
  AWorkMode: TWorkMode; AWorkCount: Integer);
begin
frmRegistration.ProgressBar1.Position:=http1.Response.ContentStream.Size;
end;

procedure TGetPrices.closeDialog;
begin
with frmRegistration do
  begin
  if problems='' then
    Close
  else
    begin
    btnCancel.Caption:='&Close';
    Memo1.Text:='Closing prices for '+Copy(problems,1,length(problems)-2)+' are not yet available.';
    end;
  end;
Form1.refreshLists;
end;

procedure TGetPrices.initProgressBar;
begin
with frmRegistration.ProgressBar1 do
  begin
  Max:=downloadSize;
  Position:=0;
  end;
end;

procedure TGetPrices.updateProgressBar;
begin
frmRegistration.ProgressBar1.StepIt;
end;

procedure TGetPrices.wrapUpMarkets;
begin
if FileExists(newMarketsFileName){ and (not FileExists(marketsFileName) or
(MessageDlg('A new list of indices has been downloaded. '
+' Would you like to replace the old list now?',mtConfirmation, [mbYes, mbNo], 0)=mrYes))} then
  begin
  DeleteFile(marketsFileName);
  RenameFile(newMarketsFileName,marketsFileName);
  Form1.processMarketsFile;
  end;
end;

procedure TGetPrices.displayMessage;
begin
frmRegistration.Memo1.Text:=msg;
end;

procedure TGetPrices.changeButton;
begin
frmRegistration.btnCancel.Caption:='&Close';
end;

procedure TGetPrices.restorePrices;
var
  marketSymbol: String;
begin
marketSymbol:=TrimRight(getSymbol(market));
if not Form1.getPricesFile(marketSymbol,market,true) and not Form1.getPricesFile(marketSymbol,market,false) then
  begin
  datesArray[market,length(datesArray[market])-1]:=0;
  pricesChanged[market]:=false;
  end;
end;

procedure TGetPrices.downloadSharePrices;
var
  mktName, yahooStr: String;
begin
http1.Free;
mktName:=TrimRight(Copy(marketSymbols,markets[market]*nameLength + 1,nameLength));
msg:='Downloading historical prices for '+mktName;
initializeHTTP;
http1.HandleRedirects:=true;
downloadSize:=1000000;
initProgressBar;
try
  yahooStr:=http1.Get('http://'+mainPricesPrefix+mktName)
except
  on EIdClosedSocket do Exit;//cancelled?
  else yahooStr:='';
end;
if yahooStr<>'' then processPrices(yahooStr,market);
end;


procedure TGetPrices.downloadMarketPrices;
var
  i,j,k,len,dateNo,shareNo,c,lines,code,newNoDays,lastDate,dateOffset,companyLen,diff:Integer;
  yahooStr,cName,mktSymbol,yahooStr2,previousStr,mktName:String;
  year,fromYear,month,day,latestTime,latestDate:Word;
  latestDateAvail,firstDateRequired:TDateTime;
  downloadDays:Integer;
  tempPricesArray:array of Single;
  tempDatesArray:array of Word;
  histThread: array[0..4] of TThread;
  latestPrice,price:Single;
  compFlags:array of Boolean;
  wronglyDelisted:Boolean;
//  latestAvailable:Boolean;//indicates that latest non-historical price is available (i.e. market has closed)
const
  numThreads:Integer=5;
label
  problemExit;
begin
if urlCountry[markets[market]]='' then //yahoo servers in USA
  begin
  latestUrl:=mainConstituentsPrefix;
  histUrl:=mainPricesPrefix;
  end
else
  begin
  latestUrl:=urlCountry[markets[market]]+nonUSsuffix+constituentsUrl;
  histUrl:=regionalPricesPrefix;
  end;
mktSymbol:=TrimRight(Copy(marketSymbols,markets[market]*nameLength + 1,nameLength));
if mktSymbol[1]<>'^' then
  begin
  mktName:=mktSymbol;
  mktSymbol:=indexPrefix+mktSymbol;
  end
else
  begin
  Delete(mktSymbol,1,1);
  mktName:=mktSymbol;
  end;
msg:='Downloading latest price for '+mktName+' index.';
//determine if today's prices are available by examining latest market time and date
http1.Free;
initializeHTTP;
http1.HandleRedirects:=true;
try
  //download current index from U.S. site: more reliable
  yahooStr:=http1.Get('http://'+mainConstituentsPrefix+mktSymbol+indexSuffix);
  SetLength(yahooStr,DeleteC(yahooStr,'"'));
except
  on EIdClosedSocket do Exit;//cancelled?
  else yahooStr:='';
end;
len:=length(yahooStr);
if (len=0) or (yahooStr[len]<>chr(10)) then goto problemExit;
i:=1;
j:=i;
while yahooStr[i]<>',' do inc(i);
latestPrice:=strToPrice(Copy(yahooStr,j,i-j));
inc(i);//start of date
readDate(yahooStr,i,year,month,day);
oldNoDays:=length(datesArray[market]);
if month>0 then
  begin
  if (oldNoDays>0) and (datesArray[market,oldNoDays-1]=dateToWord(day,month,fromYear)) then goto problemExit;
  latestDateAvail:=EncodeDate(year,month,day);
  yahooStr2:='';
  end
else //error reading date, including if date='n/a'
  begin
  //must deduce date by comparing price with historical prices
  DecodeDate(today-1,year,month,day);
  pricesFromDateUrl:='&a='+IntToStr(month-1)+'&b='+IntToStr(day)+'&c='+IntToStr(year){+'&ignore=.csv'};
  try
    yahooStr2:=http1.Get('http://'+histUrl+mktSymbol+pricesFromDateUrl);
  except
    on EIdClosedSocket do Exit;//cancelled
    else yahooStr2:='';
  end;
  len:=length(yahooStr2);
  if (len=0) or (yahooStr2[len]<>chr(10)) then
    latestDateAvail:=0//date unknown: latestAvailable will be set to false in next block
  else
    begin
    j:=len;
    Repeat dec(j) Until yahooStr2[j]=',';
    inc(j);
    //if prices differ, then one is today's and the other is previous mkt day
    if strToPrice(Copy(yahooStr2,j,len-j))=latestPrice then
      begin
      j:=1;
      readDate(yahooStr2,j,year,month,day);
      latestDateAvail:=EncodeDate(year,month,day);
      end
    else
      begin
      latestDateAvail:=today;
      DecodeDate(latestDateAvail,year,month,day);
      end;
    end;
  end;
{if latestDateAvail>=today then //same day after opening, so check closing time
  //(use >= to cater for countries such as New Zealand ahead in time)}
if (latestDateAvail>0) and (yahooStr2='') then
  begin
  //check time is after market close
  Repeat inc(i) Until (yahooStr[i]>='0') and (yahooStr[i]<='9');
  j:=i;
  Repeat inc(i) Until (yahooStr[i]<'0') or (yahooStr[i]>'9'); //allow for any separator (e.g. ':')
  k:=i;
  inc(i);
  Repeat inc(i) Until (yahooStr[i]<'0') or (yahooStr[i]>'9');
  latestTime:=StrToInt(Copy(yahooStr,j,k-j) + Copy(yahooStr,k+1,i-k-1));
  if ((yahooStr[i]='p') or (yahooStr[i]='P')) and (latestTime<1200) then inc(latestTime,1200);
  //latestAvailable prompts use of bulk download prices given in names download
  latestAvailable:=latestTime>=closingTimes[markets[market]];
  end
else
  latestAvailable:=false;
  //latestAvailable:=true;
if latestAvailable then
  begin
  latestDate:=dateToWord(day,month,year);
  //don't need to download latest date
  DecodeDate(latestDateAvail-1,year,month,day);
  pricesToDateUrl:='&d='+IntToStr(month-1)+'&e='+IntToStr(day)+'&f='+IntToStr(year)+'&g=d';
  end
else
  pricesToDateUrl:='';
//determine last downloaded date of market
if oldNoDays>0 then //data already exists
  begin
  //historical prices previously downloaded
  day:=datesArray[market,oldNoDays-1] and 31;
  month:=(datesArray[market,oldNoDays-1] shr 5) and 15;
  year:=(datesArray[market,oldNoDays-1] shr 9) + 1900;
  //increment to next day
  firstDateRequired:=EncodeDate(year,month,day)+1;
  DecodeDate(firstDateRequired,year,month,day);
  end
else
  begin
  //empty market, so download all years of prices
  SetLength(yahooPrices[market],1);
  firstDateRequired:=0;
  //index is always at position zero
  namesArray[market]:=padName('^'+mktName);
  numDelisted[market]:=0;
  end;
if not latestAvailable or (firstDateRequired<latestDateAvail) then
  begin
  //get historical market dates and closing prices for whole index
  //(may not be necessary if only one day outstanding, but can only tell after the event
  if length(datesArray[market])>0 then
    pricesFromDateUrl:='&a='+IntToStr(month-1)+'&b='+IntToStr(day)+'&c='+IntToStr(year){+'&ignore=.csv'}
  else
    pricesFromDateUrl:='';//fetch all available
  http1.Free;
  msg:='Downloading historical prices for '+mktName +' index';
  initializeHTTP;
  http1.HandleRedirects:=true;
  downloadSize:=1000000;
  initProgressBar;
  try
    yahooStr:=http1.Get('http://'+histUrl+mktSymbol+pricesToDateUrl+pricesFromDateUrl);
  except
    on EIdClosedSocket do Exit;//cancelled?
    else yahooStr:='';
  end;
  len:=length(yahooStr);
  if len>0 then
    begin
    downloadDays:=CountF(yahooStr,chr(10),1);//don't subtract 1 for header, as extra could be needed
    if not latestAvailable then dec(downloadDays);
    if (downloadDays=0) or (yahooStr[len]<>chr(10)) then goto problemExit;
    SetLength(tempDatesArray,downloadDays);
    SetLength(tempPricesArray,downloadDays);
    dateNo:=downloadDays-1;
    if latestAvailable then
      begin
      tempDatesArray[dateNo]:=latestDate;
      tempPricesArray[dateNo]:=latestPrice;
      dec(dateNo);
      end;
    i:=1;
    while yahooStr[i]<>chr(10) do inc(i);//skip header
    inc(i);
    //populate dates array (date is 1st column of each line)
    //and overall index price (closing is 5th column)
    while (dateNo>=0) and (i<len) do
      begin
      //note: hist date format is yyyy-mm-dd
      try
        year:=StrToInt(Copy(yahooStr,i,4));
        month:=StrToInt(Copy(yahooStr,i+5,2));
        day:=StrToInt(Copy(yahooStr,i+8,2));
        tempDatesArray[dateNo]:=dateToWord(day,month,year);
      except
        month:=0;
        end;
      if (month=0) then goto problemExit;
      //move to start of Close or Adj. Close
      c:=commaCount;
      inc(i,13);
      Repeat
        while (yahooStr[i]<>',') do inc(i);
        inc(i);
        dec(c);
      Until c=0;
      j:=i;
      //move to next comma or end of line
      if commaCount=3 then
        begin
        Repeat inc(i) Until yahooStr[i]=',';
        tempPricesArray[dateNo]:=strToPrice(Copy(yahooStr,j,i-j));
        while yahooStr[i]<>chr(10) do inc(i);//move to next line
        end
      else
        begin
        Repeat inc(i) Until yahooStr[i]=chr(10);
        tempPricesArray[dateNo]:=strToPrice(Copy(yahooStr,j,i-j));
        end;
      dec(dateNo);
      inc(i);
      end;
    //inc(endDateNo,oldNoDays);//for historical thread, downloading individual shares, populating yahooPrice directly
    newNoDays:=oldNoDays+downloadDays;
    SetLength(datesArray[market],newNoDays);
    SetLength(yahooPrices[market,0],newNoDays);
    if oldNoDays>0 then
      for i:=1 to numCompanies(market)-1 do
        begin
        companyLen:=length(yahooPrices[market,i]);
        SetLength(yahooPrices[market,i],companyLen+downloadDays);
        FillChar(yahooPrices[market,i,companyLen],4*downloadDays,0);
        end;
    inc(dateNo);
    Move(tempPricesArray[dateNo],yahooPrices[market,0,oldNoDays],4*downloadDays);
    Move(tempDatesArray[dateNo],datesArray[market,oldNoDays],2*downloadDays);
    for i:=oldNoDays+1 to newNoDays-2 do
      if (yahooPrices[market,0,i]=0) or
      ((yahooPrices[market,0,i-1]>2*yahooPrices[market,0,i])
      and (yahooPrices[market,0,i+1]>2*yahooPrices[market,0,i]))
      or ((yahooPrices[market,0,i]>2*yahooPrices[market,0,i-1])
      and (yahooPrices[market,0,i]>2*yahooPrices[market,0,i+1])) then
        yahooPrices[market,0,i]:=-Abs(yahooPrices[market,0,i-1]);
    end
  else if latestAvailable then
    begin
    //don't need to download historical prices
    downloadDays:=1;
//    newNoDays:=oldNoDays+1;
    SetLength(datesArray[market],oldNoDays+1);
    for i:=0 to numCompanies(market)-1 do
      begin
      companyLen:=length(yahooPrices[market,i]);
      SetLength(yahooPrices[market,i],companyLen+1);
      yahooPrices[market,i,companyLen]:=0;
      end;
    datesArray[market,oldNoDays]:=latestDate;//tempDatesArray[downloadDays-1];
//    SetLength(yahooPrices[market,0],oldNoDays+1);
    yahooPrices[market,0,oldNoDays]:=latestPrice;//tempPricesArray[downloadDays-1];
    end
  else
    goto problemExit;
  end
else if firstDateRequired=latestDateAvail then
  begin
  //don't need to download historical prices
  downloadDays:=1;
//  newNoDays:=oldNoDays+1;
  SetLength(datesArray[market],oldNoDays+1);
  for i:=0 to numCompanies(market)-1 do
    begin
    companyLen:=length(yahooPrices[market,i]);
    SetLength(yahooPrices[market,i],companyLen+1);
    yahooPrices[market,i,companyLen]:=0;
    end;
  datesArray[market,oldNoDays]:=latestDate;
//  SetLength(yahooPrices[market,0],oldNoDays+1);
  yahooPrices[market,0,oldNoDays]:=latestPrice;
  end
else
  goto problemExit;
if (indexOnly and (oldNoDays=0)) or ((length(yahooPrices[market])=1) and (oldNoDays>0)) then
  begin
  pricesChanged[market]:=true;
  Exit;//don't download constituents
  end;
pricesFromDateUrl:='&a='+IntToStr(month-1)+'&b='+IntToStr(day)+'&c='+IntToStr(year);
//get names of latest market constituents along with latest day's prices
//zero prices
c:=0;
msg:='Downloading ' + mktName + ' index constituents';
Synchronize(displayMessage);
downloadSize:=10;
initProgressBar;
http1.Free;
http1:=TIdHTTP.Create;
http1.HandleRedirects:=true;
setProxyOptions(http1);
yahooStr:='';
SetLength(compFlags,length(namesArray[market]));
wronglyDelisted:=false;
Repeat
  previousStr:=yahooStr;
  try
    yahooStr:=http1.Get('http://'+latestUrl+namePrefix+mktSymbol
    +constituentsSuffix+constituentsCounter+IntToStr(c));
    SetLength(yahooStr,DeleteC(yahooStr,'"'));
  except
    on EIdClosedSocket do Exit;//cancelled?
    else yahooStr:='';
  end;
  lines:=0;
  if yahooStr=previousStr then Break;
  len:=length(yahooStr);
  if (len=0) or (yahooStr[len]<>chr(10)) or (yahooStr[1]='@') then
    begin
    latestAvailable:=false;
    pricesToDateUrl:='';
    Break;
    end;
  //read epic codes into string of namesArray
  i:=1;
  dec(len);
  while i<len do
    begin
    j:=i;
    while (yahooStr[i]<>'/') do inc(i);
    dec(i,2);
    readDate(yahooStr,i,year,month,day);
    //read date: only use price or create new share if date is latest (i.e. exclude suspended shares)
    if (month=0) or (EncodeDate(year,month,day)=latestDateAvail) then
      begin
      k:=j;
      while yahooStr[j]<>',' do inc(j);
      cName:=Copy(yahooStr,k,j-k)+StringOfChar(' ',nameLength-(j-k));
      if length(cName)<=nameLength then
        begin
        shareNo:=Pos(cName,namesArray[market]);
        if (shareNo=0) then
          begin
          //new share
          shareNo:=numCompanies(market);
          namesArray[market]:=namesArray[market]+cName;
          SetLength(yahooPrices[market],shareNo+1);//first share is index price
          SetLength(yahooPrices[market,shareNo],1);
          SetLength(compFlags,shareNo+1);
          end
        else
          begin
          shareNo:=(shareNo-1) div nameLength;
          if shareNo<numDelisted[market] then wronglyDelisted:=true;
          end;
        if latestAvailable and not compFlags[shareNo] then
          begin
          inc(j);
          k:=j;
          while (yahooStr[k]<>',') do inc(k);
//          lastDate:=length(yahooPrices[market,shareNo]);
//          SetLength(yahooPrices[market,shareNo],lastDate+1);
          yahooPrices[market,shareNo,length(yahooPrices[market,shareNo])-1]:=strToPrice(Copy(yahooStr,j,k-j));
          compFlags[shareNo]:=true;
          end;
        end;
      end;
    while yahooStr[i]<>chr(10) do inc(i);
    inc(i);
    inc(lines);
    end;
  inc(c,constituentsInc);
  updateProgressBar;
Until (lines<>51);
if length(yahooPrices[market])>1 then
  begin
  noCompanies:=numCompanies(market);
  if (downloadDays>1) or (not latestAvailable) or wronglyDelisted then
    begin
    if wronglyDelisted then
      globalShareNo:=0
    else
      globalShareNo:=numDelisted[market];
    downloadSize:=noCompanies-1;//exclude market symbol
    Synchronize(initProgressBar);
    msg:='Downloading '+mktName +' constituent prices';
    Synchronize(displayMessage);
    SetLength(extraDates,0);
    //create threads to download historical prices for individual shares
    InitializeCriticalSection(CS1);
    InitializeCriticalSection(CS2);
    threadsRunning:=numThreads;
    for i:=0 to numThreads-1 do histThread[i]:=TGetHistorical.Create(false);
    Repeat
      if Terminated then
        for i:=0 to numThreads-1 do
          begin
          histThread[i].Terminate;
          histThread[i].WaitFor;
          end;
    Until threadsRunning=0;
    DeleteCriticalSection(CS1);
    DeleteCriticalSection(CS2);
    //if majority of shares have an extra date, then apply to all shares
    len:=length(datesArray[market]);
    for i:=0 to length(extraDates)-1 do
      begin
      k:=0;
      for j:=0 to noCompanies-1 do if extraPrices[i,j]>0 then inc(k);
      if k>noCompanies div 2 then //more than half have extra date
        begin
        if datesArray[market,0]<extraDates[i] then
          begin
          dateNo:=len;
          Repeat dec(dateNo) Until datesArray[market,dateNo]<extraDates[i];
          inc(dateNo);
          end
        else
          dateNo:=0;
        SetLength(datesArray[market],len+1);
        Move(datesArray[market,dateNo],datesArray[market,dateNo+1],(len-dateNo)*2);
        datesArray[market,dateNo]:=extraDates[i];
        diff:=len-dateNo;
        for j:=0 to noCompanies-1 do
          begin
          dateOffset:=length(yahooPrices[market,j]);
          if (dateOffset>diff) or ((dateOffset=diff) and (extraPrices[i,j]>0)) then
            begin
            SetLength(yahooPrices[market,j],dateOffset+1);
            dec(dateOffset,diff);
            Move(yahooPrices[market,j,dateOffset],yahooPrices[market,j,dateOffset+1],diff*4);
            if extraPrices[i,j]>0 then
              yahooPrices[market,j,dateOffset]:=extraPrices[i,j]
            else
              yahooPrices[market,j,dateOffset]:=-yahooPrices[market,j,dateOffset-1];
            end;
          end;
        inc(len);
        end;
      end;
    end;
  processDelistedShares(market);
  //sometimes most recent price is quoted in different units, so compensate when necessary
  for i:=numDelisted[market]+1 to numCompanies(market)-1 do
    begin
    lastDate:=length(yahooPrices[market,i])-1;
    if (yahooPrices[market,i,lastDate]>0) and (lastDate>0) and (yahooPrices[market,i,lastDate-1]>0)
    and(yahooPrices[market,i,lastDate]/yahooPrices[market,i,lastDate-1]<0.2) then
      yahooPrices[market,i,lastDate]:=100*yahooPrices[market,i,lastDate];
    end;
  //fill any gaps with negative of preceding day's price
  for i:=0 to noCompanies-1 do
    begin
    for j:=1 to length(yahooPrices[market,i])-1 do
      if yahooPrices[market,i,j]=0 then yahooPrices[market,i,j]:=-Abs(yahooPrices[market,i,j-1]);
    end;
  end;
pricesChanged[market]:=true;
Exit;
problemExit: problems:=problems+mktName+', ';
end;

procedure readDate(const yahooStr:String; var i:Integer; out year,month,day:Word);
var
  m:Integer;
begin
m:=i;
//see if date is missing
while (yahooStr[m]<>'/') and (yahooStr[m]<>',') do inc(m);
if yahooStr[m]=',' then
  month:=0
else
  try
  //get date
  m:=i;//start of date
  Repeat inc(i) Until yahooStr[i]='/';
  month:=StrToInt(Copy(yahooStr,m,i-m));
  inc(i);
  m:=i;
  Repeat inc(i) Until yahooStr[i]='/';
  day:=StrToInt(Copy(yahooStr,m,i-m));
  inc(i);
  m:=i;
  Repeat inc(i) Until (yahooStr[i]<'0') or (yahooStr[i]>'9');
  year:=StrToInt(Copy(yahooStr,m,i-m));
  except
    month:=0;
  end;
end;

end.
