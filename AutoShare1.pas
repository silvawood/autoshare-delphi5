unit AutoShare1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Spin, ExtCtrls, ComCtrls, Grids, Math, clipbrd,
  registry, Nag, Buttons, Menus, hyperStr, ScktComp,
  OleCtrls, Gauges, Printers, CheckLst, Zlib, HH, HH_FUNCS, ShellAPI, jpeg;

type
  TGraphArray = array of array[0..2] of Single;
  TTabListBox = class(TListBox)
  public
    procedure CreateParams(var Params: TCreateParams); override;
  end;
  TForm1 = class(TForm)
    lstAllShares: TListBox;
    lblAllShares: TLabel;
    btnSave: TButton;
    btnHelp: TButton;
    btnExit: TButton;
    btnUndo: TButton;
    PageControl2: TPageControl;
    averagesSheet: TTabSheet;
    moversSheet: TTabSheet;
    delistedSheet: TTabSheet;
    lblMovingDown: TLabel;
    lblMovingUp: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    lstMovingDown: TListBox;
    lstMovingUp: TListBox;
    spnAverage1: TSpinEdit;
    spnAverage2: TSpinEdit;
    lblDelisted: TLabel;
    lstDelisted: TListBox;
    btnDelete: TButton;
    btnCombine: TButton;
    btnSelectAll: TButton;
    lblRisers: TLabel;
    lblFallers: TLabel;
    pickerFrom: TDateTimePicker;
    pickerTo: TDateTimePicker;
    Label5: TLabel;
    Label7: TLabel;
    comboPeriod: TComboBox;
    Label8: TLabel;
    chkRecentCrossing: TCheckBox;
    btnUpdate: TButton;
    spnCrossing: TSpinEdit;
    lblCrossing: TLabel;
    portfolioSheet: TTabSheet;
    btnRemove: TButton;
    grdPortfolio: TStringGrid;
    btnAdd: TButton;
    lblRecentList: TLabel;
    lstRecentIssues: TListBox;
    spnRecentDays: TSpinEdit;
    lblRecentDays2: TLabel;
    lblAvg2: TLabel;
    chkShowGrid: TCheckBox;
    lblAvg1: TLabel;
    grdPrice: TStringGrid;
    lblGraphTitle: TLabel;
    highLowSheet: TTabSheet;
    spnHighLow: TSpinEdit;
    Label1: TLabel;
    lstNearHigh: TListBox;
    lstNearLow: TListBox;
    lblHighs: TLabel;
    lblLows: TLabel;
    Label9: TLabel;
    Graph: TPaintBox;
    chkRising: TCheckBox;
    chkFalling: TCheckBox;
    btnHint: TSpeedButton;
    PopupMenu1: TPopupMenu;
    AddtoPortfolio: TMenuItem;
    lblRange: TLabel;
    DateTimePicker1: TDateTimePicker;
    lblSearch: TLabel;
    edtSearch: TEdit;
    lblTotal: TLabel;
    optSimple: TRadioButton;
    optExp: TRadioButton;
    Label6: TLabel;
    ColorDialog1: TColorDialog;
    ColourMenu: TPopupMenu;
    ChangeColour: TMenuItem;
    Returntodefault: TMenuItem;
    Copytoclipboard1: TMenuItem;
    Label10: TLabel;
    btnPortfolioUndo: TButton;
    backtestSheet: TTabSheet;
    Label11: TLabel;
    spnTestAv1: TSpinEdit;
    Label13: TLabel;
    spnTestAv2: TSpinEdit;
    optTestSimple: TRadioButton;
    optTestExp: TRadioButton;
    btnGetResults: TButton;
    lblResults: TLabel;
    Memo1: TMemo;
    Image1: TImage;
    lblContourTip: TLabel;
    Label12: TLabel;
    btnResize: TButton;
    popupPortfolio: TPopupMenu;
    printPortfolio: TMenuItem;
    exportPortfolio: TMenuItem;
    SaveDialog1: TSaveDialog;
    chkClearResults: TCheckBox;
    lstPortfolio: TComboBox;
    lblPortfolio: TLabel;
    spnMargin: TSpinEdit;
    spnMaxBet: TSpinEdit;
    lblMaxBet: TLabel;
    chkSpreadBet: TCheckBox;
    lstRange: TComboBox;
    lstHighLowRange: TComboBox;
    txtMinBet: TEdit;
    TabSheet1: TTabSheet;
    chkScope: TCheckBox;
    lstScope: TComboBox;
    spnThreshold: TSpinEdit;
    lstMarkets: TCheckListBox;
    Label2: TLabel;
    chkSymbol: TCheckBox;
    lstSymbol: TComboBox;
    lblSymbol: TLabel;
    lblIMR: TLabel;
    btnImport: TButton;
    OpenDialog1: TOpenDialog;
    btnExport: TButton;
    ShowFundamentals1: TMenuItem;
    lblSpread: TLabel;
    spnSpread: TSpinEdit;
    lblMinBet: TLabel;
    chkUnadjusted: TCheckBox;
    popupGraph: TPopupMenu;
    printGraph: TMenuItem;
    saveGraph: TMenuItem;
    copyGraph: TMenuItem;
    PrintDialog1: TPrintDialog;
    chkIndexOnly: TCheckBox;
    popupEditPrice: TPopupMenu;
    Addday1: TMenuItem;
    graphScroll: TScrollBar;
    chkListTrades: TCheckBox;
    chkOptimize: TCheckBox;
    lblRecentDays1: TLabel;
    lstTradeType: TComboBox;
    lblTradeType: TLabel;
    btnClear: TButton;
    Editday1: TMenuItem;
    showPrice: TMenuItem;
    Inserttoday1: TMenuItem;
    resetAvgValues: TMenuItem;
    D1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure lstMovingDownClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure lstAllSharesClick(Sender: TObject);
    procedure lstMovingUpClick(Sender: TObject);
    procedure lstDelistedClick(Sender: TObject);
    procedure lstAllSharesMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
    procedure spnAverage1Change(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnCombineClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnHelpClick(Sender: TObject);
    procedure btnUndoClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure btnSelectAllClick(Sender: TObject);
    procedure lstRisersClick(Sender: TObject);
    procedure lstFallersClick(Sender: TObject);
    procedure autoSelectAverageItem;
    procedure refreshLists;
    procedure calculateGraph;
    procedure calcAverage(spinSetting: Integer; arrayIndex: Integer);
	  procedure populateUpDownLists;
    procedure autoSelectMoversItem;
    procedure updateMoversLists;
    procedure comboPeriodChange(Sender: TObject);
    procedure pickerToCloseUp(Sender: TObject);
    procedure pickerFromCloseUp(Sender: TObject);
    procedure averagesSheetShow(Sender: TObject);
    procedure moversSheetShow(Sender: TObject);
    procedure updateRightControls(const name:String);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure chkRisingAverageClick(Sender: TObject);
    procedure chkRecentCrossingClick(Sender: TObject);
    procedure chkShowGridClick(Sender: TObject);
    procedure btnUpdateClick(Sender: TObject);
    procedure chkHighLowClick(Sender: TObject);
    procedure spnCrossingChange(Sender: TObject);
    procedure lstBoughtDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure btnRemoveClick(Sender: TObject);
    procedure grdPortfolioDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure portfolioSheetShow(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure addPortfolioShare(rowNo:Integer);
    procedure removePortfolioShare(r:Integer);
    procedure grdPortfolioKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure grdPortfolioSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure grdPortfolioGetEditText(Sender: TObject; ACol, ARow: Integer;
      var Value: String);
    procedure autoSelectPortfolioItem;
	  function processSpinEdit(var spinControl:TSpinEdit):Boolean;
    procedure spnAverage2Change(Sender: TObject);
    procedure btnHelpMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure updateRecentIssues;
    procedure spnRecentDaysChange(Sender: TObject);
    procedure lstRecentIssuesClick(Sender: TObject);
    procedure populatePriceGrid;
    procedure grdPriceTopLeftChanged(Sender: TObject);
    procedure grdPriceSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure grdPriceKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure spnHighLowChange(Sender: TObject);
    procedure updateHighLowLists;
    procedure lstNearHighClick(Sender: TObject);
    procedure lstNearLowClick(Sender: TObject);
    procedure highLowSheetShow(Sender: TObject);
    procedure autoSelectHighLowItem;
    procedure PageControl2MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure updatePortfolioColumns(gridRow,market,nameNo:Integer);
    procedure updatePortfolioGrid;
    procedure plotGraph(toScreen:Boolean);
    procedure GraphPaint(Sender: TObject);
    procedure refreshGraph;
    procedure refreshGraphSelection;
    procedure GraphMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure GraphMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure changePortfolioPrice;
    procedure GraphMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure showGraphTip(X,Y:Integer; Shift:TShiftState);
    function getPricesFile(const fileName:String; market:Integer; getOriginal: Boolean):Boolean;
    function getPortfolioFile(getOriginal: Boolean):Boolean;
    procedure chkRisingClick(Sender: TObject);
    procedure btnHintClick(Sender: TObject);
    procedure AddtoPortfolioClick(Sender: TObject);
    procedure grdPortfolioColumnMoved(Sender: TObject; FromIndex,
      ToIndex: Integer);
    procedure grdPortfolioDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure MonthCalendar1Click(Sender: TObject);
    procedure grdPortfolioTopLeftChanged(Sender: TObject);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure grdPortfolioRowMoved(Sender: TObject; FromIndex,
      ToIndex: Integer);
    procedure grdPortfolioMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grdPortfolioDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure grdPortfolioExit(Sender: TObject);
    procedure updatePortfolioTotal;
    procedure optSimpleClick(Sender: TObject);
    procedure simpleAveragePopulate;
    procedure exponentialAveragePopulate;
    procedure changeSpinValues(averagesTab:Boolean);
    procedure savePortfolio;
    procedure ChangeColourClick(Sender: TObject);
    procedure ReturntodefaultClick(Sender: TObject);
    procedure Copytoclipboard1Click(Sender: TObject);
    procedure btnPortfolioUndoClick(Sender: TObject);
    procedure btnClearClick(Sender: TObject);
    procedure btnGetResultsClick(Sender: TObject);
    procedure spnTestAv1Change(Sender: TObject);
    procedure spnTestAv2Change(Sender: TObject);
    procedure optExpClick(Sender: TObject);
    procedure optTestExpClick(Sender: TObject);
    function getResults(updateMemo:Boolean):Single;
    procedure Image1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Image1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Image1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure updateGauge(X,Y:Integer);
    procedure optTestSimpleClick(Sender: TObject);
    procedure findProfits;
    procedure copyPortfolioRow(fromRow,toRow:Integer);
    procedure grdPortfolioMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    function isCellLower(colType:Integer; const cellStr,compStr:String):Boolean;
    procedure spnThresholdChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure portfolioSheetMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure PageControl2Change(Sender: TObject);
    procedure btnResizeClick(Sender: TObject);
    procedure normalPortfolioSize;
    function hintDown(Sender:TObject):Boolean;
    function hintDown2(helpId:THelpContext):Boolean;
    procedure lstNearLowMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure printPortfolioClick(Sender: TObject);
    procedure exportPortfolioClick(Sender: TObject);
    procedure printHeadings(pColWidths: array of Integer);
    procedure lstPortfolioChange(Sender: TObject);
//    procedure backtestUsingSettings;
//    procedure loadYahooData;
//    procedure stochasticBacktest;
    procedure lstPortfolioKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure lstRangeChange(Sender: TObject);
    procedure lstHighLowRangeChange(Sender: TObject);
    procedure chkScopeClick(Sender: TObject);
    procedure lstScopeChange(Sender: TObject);
    procedure processMarketsFile;
    procedure lstMarketsClickCheck(Sender: TObject);
    function marketTicked(marketNo:Integer):Boolean;
    procedure lstSymbolChange(Sender: TObject);
    procedure chkSymbolClick(Sender: TObject);
    procedure removeFromPortfolio(const shareName:String);
    procedure positionDatePicker;
    procedure FormDestroy(Sender: TObject);
    procedure FormClick(Sender: TObject);
    procedure lstMarketsClick(Sender: TObject);
    procedure chkSpreadBetClick(Sender: TObject);
    procedure chkClearResultsClick(Sender: TObject);
    procedure btnImportClick(Sender: TObject);
    procedure ShowFundamentals1Click(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure grdPriceGetEditText(Sender: TObject; ACol, ARow: Integer;
      var Value: String);
    procedure spnMaxBetClick(Sender: TObject);
    procedure spnTestAv1Click(Sender: TObject);
    procedure spnAverage1Click(Sender: TObject);
    procedure optTestExpMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure lstPortfolioExit(Sender: TObject);
    procedure editPortfolioName;
    procedure grdPortfolioEnter(Sender: TObject);
    procedure printGraphClick(Sender: TObject);
    procedure saveGraphClick(Sender: TObject);
    procedure copyGraphClick(Sender: TObject);
    procedure graphToBitmap(out bitmap:TBitmap);
    procedure edtSearchChange(Sender: TObject);
    procedure newPortfolio;
    procedure grdPortfolioMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grdPortfolioSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure deleteDayClick(Sender: TObject);
    procedure Addday1Click(Sender: TObject);
    procedure graphScrollChange(Sender: TObject);
    procedure chkListTradesClick(Sender: TObject);
    procedure chkOptimizeClick(Sender: TObject);
    procedure Editday1Click(Sender: TObject);
    procedure showPriceClick(Sender: TObject);
    procedure cancelBacktest;
    procedure lstTradeTypeChange(Sender: TObject);
    procedure Inserttoday1Click(Sender: TObject);
    procedure grdPortfolioContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
    procedure resetAvgValuesClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  numDelisted: array of Word;
  companyNo: Integer;
  namesArray: array of String;
  yahooPrices: array of array of array of Single; //up to 1000 days for each company
  datesArray: array of array of Word;
  closingTimes: array of Word;
  editValue: String;
  low, high: Single;
  graphData: TGraphArray;
  answer: Word;
  selectedName: String;
  editRow:Integer;
  oldEditValue: String;
  orderId: Integer;
  portfolioCol: array[0..10] of Integer;
  threshold:Integer;
  lowAverage,highAverage:Integer;
  factor1, factor2, factor1min1, factor2min1:Single;
  portfolioChanged, simpleAvg, expired: Boolean;
  pricesChanged: array of Boolean; 
  regKey:String;
  contourData:array[0..224] of array[1..224] of Single;
  editPrice:String;
  editedIndex:Integer;
  spreadFile:String;
  scopeSetting,currMarket,numMarkets:Integer;
  markets: array of Integer;//corresponding lstMarkets ItemIndex for each ticked market
  urlCountry: array of String;
  mainPricesPrefix,regionalPricesPrefix,constituentsUrl,mainConstituentsPrefix,constituentsSuffix,indexPrefix: String;
  indexSuffix,fundamentalsUrl,nonUSsuffix,constituentsCounter,namePrefix:String;
  constituentsInc:Integer;
  marketSymbols:String;
  addSymbol,prefix: Boolean;
  mktSymbol:String;
  lstRisers,lstFallers:TTabListBox;
  marketNotUTD: array of Boolean;
  today:TDateTime;
  mainUrl:String;
  mHHelp: THookHelpSystem;
  helpWindow: hwnd;
  sharesListed,reverseSort,portfolioMouseDown:Boolean;
  priceGridDay:Integer;
  selectionStart,selectionEnd:Integer;
  boxBitmap:TBitmap;
  boxLeft,boxTop:Integer;
  mousePressed:Boolean;
  zeroY: Boolean;

Const
  nameLength: Integer=9;
//  lastCol: Integer=10;
  fractions: String='¼½¾';
  graphLeftMargin:Integer=48;
  graphRightMargin:Integer=10;
  graphTopMargin:Integer=10;
  graphBottomMargin:Integer=17;
  months: String='JanFebMarAprMayJunJulAugSepOctNovDec';
  yIntervals: array [0..18] of Single = (0.25,0.5,1,2,5,10,20,50,100,200,500
  ,1000,2000,5000,10000,20000,50000,100000,200000);
  rangeOptions: array [0..6] of Integer=(21,63,126,252,504,1260,2520);
  marketsFileName:String = 'indices.lst';
  newMarketsFileName: String = 'newIndices.lst';
  notAvailStr:String='n/a';
  defaultStr:String='Default';
  urlSource:String='http://www.silvawood.co.uk/source.htm';
  defaultUrl:String='http://ccgi.silvawood.force9.co.uk';
  symbolSeparator:String=' ';
  columnWidths:array[0..10] of Byte=(80,61,29,21,32,32,50,35,36,36,45);
  columnLabels:array[0..10] of String=('Name of Share','Change on Day','Total %Ch',
  'Avg Pos','Price Now','Price Paid','Date of Purchase','No. Held','Value Paid','Value Now','Moving Avgs');
  newPortfolioText:String='[Add new portfolio]';

  function cellValue(const editValue:String):Single;
  function companyAndMarketNo(out market:Integer; companyName:String):Integer;
//  function companyNumber(const companyName:String):Integer;
  function companyName(market,company:Integer):String;
  function withinScope(market,company:Integer):Boolean;
  function reverseMoverStr(const moverStr:String):String;
//  function priceWithFraction(p: Integer):String;overload;
//  function priceWithFraction(p: Single):String;overload;
  procedure rangeCheck(var varToCheck:Integer;lowerLimit,upperLimit:Integer);
  procedure deleteShare(market,shareNo:Integer);
  function dateStrToWord(const S:String):Word;
  function twoDigits(dateComponent:Word):String;
  function getFirstDay(average:Integer):Integer;
  function calcExpAverage(market,company:Integer; out average1, average2: Single):Boolean;
  function percentStr(i:Single):String;
  function dayToStr(day:Integer):String;overload;
  function dayToStr(market, day:Integer):String;overload;
  function daysLabel(singular:Boolean):String;
  procedure portfolioAltered;
  function colorToCode(colorValue:TColor):Byte;
  function colorFromCode(colorCode:Byte):TColor;
  function duplicateExists(market:Integer; companyName:String):Boolean;
  function percentChange(const changeStr:String):Integer;
  function wordToPrice(w:Word):Single;
  function priceToWord(p:Single):Word;
  procedure allocMarketSpace(n:Integer);
  function fetchLine(var P:PChar):String;
  function anyPricesChanged:Boolean;
  function mktArrayIndex(i:Integer):Integer;
  function numCompanies(marketNo:Integer):Integer;
  procedure setSymbol(market:Integer);
  function padName(const name:String):String;
  function getSymbol(marketNo:Integer):String;
  function getName(marketNo,company:Integer):String;
  procedure extractNameAndSymbol(const S:String; out companyName,symbol:String);
  function priceToStr(p:Single):String;
  function strToPrice(const priceStr:String):Single;
  function dateToWord(day,month,year:Word):Word;
  procedure processDelistedShares(market:Integer);
  procedure errorMessage(const S:String);
  procedure processPrices(const S:String; market:Integer);
  function belowAverage(noDays,market,nameNo:Integer; simpleAvgType:Boolean; lowAvg,highAvg:Word):Boolean;
  procedure updateGraphTitle(noDays:Integer);
  procedure insertDay(day,month,year:Word);

{ function priceValue(const priceString: String):Integer;
  procedure savePrice(const nameString,priceString:String; today:Integer);
  procedure changePrice(const nameString,changeString:String; today:Integer);
  function stripSymbol(const S:String):String; }

implementation

uses RegistrationForm, Proxy;

{$R *.DFM}

procedure TTabListBox.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
with Params do Style := Style or LBS_USETABSTOPS;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  flags,defaults,serialnum,xy,colours,backTestOptions,betoptions:Integer;
//  company,day,count:Integer;
//  S:String;
//  newPrice:Single;
  dummy:Cardinal;
  searchResult: TSearchRec;
  name:String;
//  dayStr: String;
//  totalDiffs,leastDiff:Single;
//  F: TextFile;
//  bestMatch:String;
//  epicNo:Integer;
var
  tabStops:array[0..0] of Integer;
begin
if not SetCurrentDir('Data') then
  begin
  errorMessage('No AutoShare data could be found.');
  MkDir('Data');
  end;
mHHelp:=THookHelpSystem.Create('AutoShare.chm', '', htHHAPI);
lstRisers:=TTabListBox.Create(Self);
lstRisers.OnClick:=lstRisersClick;
lstRisers.OnMouseDown:=lstNearLowMouseDown;
lstRisers.PopupMenu:=PopupMenu1;
lstFallers:=TTabListBox.Create(Self);
lstFallers.OnClick:=lstFallersClick;
lstFallers.OnMouseDown:=lstNearLowMouseDown;
lstFallers.PopupMenu:=PopupMenu1;
portfolioMouseDown:=false;
mousePressed:=false;
tabStops[0]:=58;
with lstRisers do
  begin
  Parent:=moversSheet;
  SetBounds(44, 88, 152, 100);  // (Left, Top, Width, Height)
  SendMessage(lstRisers.Handle,LB_SETTABSTOPS, 1, LongInt(@tabStops));
  end;
with lstFallers do
  begin
  Parent:=moversSheet;
  SetBounds(44, 220, 152, 100);
  SendMessage(lstFallers.Handle,LB_SETTABSTOPS, 1, LongInt(@tabStops));
  end;
DoubleBuffered:=true;
grdPrice.DefaultColWidth:=grdPrice.Canvas.TextWidth('22-Dec-22')+5;
//populate Markets check list box, and tick markets whose source file read-only is reset
{if FileExists(newMarketsFileName) and (MessageDlg('A new list of indices has been downloaded. '
+' Would you like to replace the old list now?',mtConfirmation, [mbYes, mbNo], 0)=mrYes) then
  begin
  DeleteFile(marketsFileName);
  RenameFile(newMarketsFileName,marketsFileName);
  end;}
reverseSort:=false;
portfolioChanged:=false;
//dayRange:=numDays;
//grdPrice.ColCount:=numDays;
pickerFrom.Date:=Date-1;
pickerTo.Date:=Date;
comboPeriod.ItemIndex:=1;
lstRange.ItemIndex:=5;
lstHighLowRange.ItemIndex:=4;
lstScope.ItemIndex:=0;
lstSymbol.ItemIndex:=0;
lstTradeType.ItemIndex:=0;
expired:=false;
mainUrl:=defaultUrl;
betOptions:=0;

regKey:='Software\Silvawood\'+Application.Title;
with TRegistry.Create do
  begin
	try
		RootKey:=HKEY_CURRENT_USER;
		if OpenKey(regKey,false) then
			begin
  		GetVolumeInformation(PChar(Copy(Application.ExeName,1,3)), nil, 0, @serialnum, dummy, dummy, nil, 0);
      defaults:=ReadInteger('defaults');
    	orderId:=ReadInteger('user');//bits 0-15: userID, bits 16-31: PIN
    	flags:=ReadInteger('flags');
    	xy:=ReadInteger('xy');
    	if (defaults and 1)>0 then
    		Form1.WindowState:=wsMaximized
    	else
    		begin
      	Form1.Width:=xy and 65535;
    		Form1.Height:=xy shr 16;
      	end;
    	spnAverage1.Value:=(defaults shr 1) and 511;
    	spnAverage2.Value:=(defaults shr 10) and 511;
//      changeSpinValues(true);
      spnCrossing.Value:=(defaults shr 19) and 127;
    	chkRecentCrossing.Checked:=(flags and 32)>0;
    	lstScope.ItemIndex:=flags and 1;
      chkScope.Checked:=(flags and 2)>0;
      scopeSetting:=flags and 3; //was and 7
      chkIndexOnly.Checked:=(flags and 4)>0;
      chkRising.Checked:=(flags and 8)>0;
    	chkFalling.Checked:=(flags and 16)>0;
      if (flags and 128)>0 then optSimple.Checked:=true else optExp.Checked:=true;
    	spnHighLow.Value:=(flags shr 8) and 63;
      lstRange.ItemIndex:=(defaults shr 26) and 7;
    	chkShowGrid.Checked:=(flags and 64)>0;
      if ValueExists('colours') then
        begin
        colours:=ReadInteger('colours');
        lblAvg1.Font.Color:=colorFromCode(colours and 63);
        lblAvg2.Font.Color:=colorFromCode(colours shr 6 and 63);
        lstHighLowRange.ItemIndex:=colours shr 12 and 7;
        spnThreshold.Value:=colours shr 18;
        end;
      if ValueExists('backtest') then
        begin
        backTestOptions:=ReadInteger('backtest');
        if (backTestOptions and 1)>0 then
          optTestSimple.Checked:=true
        else
          optTestExp.Checked:=true;
        chkOptimize.Checked:=(backTestOptions and 2)>0;
        chkListTrades.Checked:=(backTestOptions and 4)>0;
        chkSpreadBet.Checked:=(backTestOptions and 8)>0;
//        chkSpreadBetClick(Sender);
        chkClearResults.Checked:=(backTestOptions and 16)>0;
        chkUnadjusted.Checked:=(backTestOptions and 32)>0;
        lstTradeType.ItemIndex:=(backTestOptions shr 6) and 3;
    	  spnTestAv1.Value:=(backTestOptions shr 8) and 511;
    	  spnTestAv2.Value:=(backTestOptions shr 17) and 511;
        betOptions:=ReadInteger('betoptions');
        spnMargin.Value:=(betoptions and 16383);
        spnMaxBet.Value:=(betoptions shr 16) and 511;
        chkSymbol.Checked:=(betOptions shr 25 and 1)>0;
        lstSymbol.ItemIndex:=betOptions shr 26 and 1;
        txtMinBet.Text:=ReadString('minbet');
        betOptions:=ReadInteger('betoptions2');
        spnSpread.Value:=(betoptions and 1023);
        end;
      if ValueExists('url') then mainUrl:=ReadString('url');
    	if (orderId and 131071)<>(((flags shr 14) xor serialnum xor xy xor defaults) and 131071) then orderId:=0;
      end
  	else
  		orderId:=0;
	except
  	orderId:=0;
  	end;
	Free;
  end;
// orderId:=0;
zeroY:=false;
processMarketsFile;
addSymbol:=chkSymbol.Checked;
prefix:=lstSymbol.ItemIndex=0;
changeSpinValues(true);
threshold:=spnThreshold.Value;
//populate lstPortfolio
if FindFirst({ExtractFileDir(Application.ExeName)+'\}'*.dat', faAnyFile, searchResult) = 0 then
  Repeat
    if searchResult.Size<10000 then
      begin
      name:=Copy(searchResult.Name,1,length(searchResult.Name)-4);
      lstPortfolio.Items.Add(name);
      end;
  Until FindNext(searchResult) <> 0;
FindClose(searchResult);
with lstPortfolio do
  begin
  Items.Add(newPortfolioText);
  editedIndex:=betOptions shr 10;
  if editedIndex=Items.Count-1 then editedIndex:=0;
  ItemIndex:=editedIndex;
  end;
updatePortfolioGrid;
refreshLists;
lstRangeChange(nil);
if orderId<>0 then Application.ShowHint:=false;
end;

procedure allocMarketSpace(n:Integer);
begin
SetLength(namesArray,n);
SetLength(yahooPrices,n);//dummy in the case of tt prices
SetLength(datesArray,n);//,numDays);
SetLength(numDelisted,n);
SetLength(markets,n);
SetLength(pricesChanged,n);
end;

function colorFromCode(colorCode:Byte):TColor;
var
  bitSet, colorByte: Integer;
begin
Result:=0;
for bitSet:=1 to 3 do
  begin
  Case colorCode and 3 of
    0: colorByte:=0;
    1: colorByte:=128;
    2: colorByte:=192;
  else
    colorByte:=255;
    end;
  Result:=(Result shl 8) or colorByte;
  colorCode:=colorCode shr 2;//get next bits
  end;
end;

function colorToCode(colorValue:TColor):Byte;
var
  byteNo, colorBits: Byte;
begin
Result:=0;
for byteNo:=1 to 3 do
  begin
  Case colorValue and 255 of
    0: colorBits:=0;
    128: colorBits:=1;
    192: colorBits:=2;
  else
    colorBits:=3;
    end;
  Result:=(Result shl 2) or colorBits;
  colorValue:=colorValue shr 8;//get next byte
  end;
end;

function TForm1.getPricesFile(const fileName:String; market:Integer; getOriginal: Boolean):Boolean;
var
  i:Integer;
  noCompanies,noDays:Word;
  extension: String;
begin
if getOriginal then extension:='mkt' else extension:='bak';
Result:=FileExists(fileName+'.'+extension);
if Result then with TFileStream.Create(fileName+'.'+extension,fmOpenRead or fmShareDenyNone) do
  try
	namesArray[market]:='';
  Read(noCompanies,2);
  Read(numDelisted[market],2);
  Read(noDays,2);
  SetLength(namesArray[market],noCompanies*nameLength);
  Read(Pointer(namesArray[market])^,noCompanies*nameLength);
  SetLength(datesArray[market],noDays);
  Read(Pointer(datesArray[market])^,noDays*2);
  SetLength(yahooPrices[market],noCompanies);
{  for i:=noDays-16 to noDays-1 do
    datesArray[market,i]:=dateToWord(datesArray[market,i] and 31, (datesArray[market,i] shr 5) and 15,
    2000+(datesArray[market,i] shr 9));}
  for i:=0 to noCompanies-1 do
    begin
    Read(noDays,2);
    SetLength(yahooPrices[market,i],noDays);
    Read(Pointer(yahooPrices[market,i])^,4*noDays);
    end;
  Free;
  pricesChanged[market]:=(not getOriginal);
  btnSave.Enabled:=portfolioChanged or anyPricesChanged;
  //update delisted list
  except
    Free;
    Result:=false;
	end;
end;

function anyPricesChanged:Boolean;
var
  i:Integer;
begin
Result:=false;
for i:=0 to numMarkets-1 do if pricesChanged[i] then Result:=true;
end;

function TForm1.getPortfolioFile(getOriginal: Boolean):Boolean;
var
  i,numShares,company,marketNo,colOrder,averages:Integer;
  purchasePrice:Single;
  purchaseDate:Word;
  fileName:String;
  companyName,marketSymbol,S: String;
  newStyle:Boolean;
  columnWidth:Byte;
  lastCol:Byte;
begin
with grdPortfolio do begin
fileName:=lstPortfolio.Text;
if getOriginal then fileName:=fileName+'.dat' else fileName:=fileName+'.bak';
Result:=FileExists(fileName);
if not Result then
  newPortfolio
else with TFileStream.Create(fileName,fmOpenRead or fmShareDenyNone) do
  begin
  RowCount:=2;
  try//except
  SetLength(companyName,nameLength);
  SetLength(marketSymbol,nameLength);
  //read colOrder details
  Read(colOrder,4);
  for i:=0 to 7 do
    portfolioCol[i]:=colOrder shr (i shl 2) and 15;
  Read(colOrder,1);
  colOrder:=colOrder and 255;
  newStyle:=(colOrder>0);
  //for newstyle, new "averages values" column added
  if newStyle then
    begin
    Read(colOrder,2);
    for i:=0 to 2 do
      portfolioCol[i+8]:=colOrder shr (i shl 2) and 15;
    lastCol:=10;
    end
  else
    begin
    Read(colOrder,1);
    for i:=0 to 1 do
      portfolioCol[i+8]:=colOrder shr (i shl 2) and 15;
    lastCol:=9;
    portfolioCol[10]:=10;
    end;
  for i:=0 to lastCol do
    begin
    Read(columnWidth,1);
    ColWidths[portfolioCol[i]]:=columnWidth;
    end;
  while Position<Size do //read share info in portfolio
    begin
 	  Read(Pointer(companyName)^,nameLength);
    company:=0;
    Read(Pointer(marketSymbol)^,nameLength);
    marketNo:=Pos(marketSymbol,marketSymbols)-1;
    if (marketNo>=0) then
      begin
      marketNo:=marketNo div nameLength;
      if lstMarkets.Checked[marketNo] then
        begin
        marketNo:=mktArrayIndex(marketNo);
        company:=Pos(companyName,namesArray[marketNo]);
        end
      else
        marketNo:=-1;
      end;
    if prefix then
      Cells[portfolioCol[0],RowCount-1]:=TrimRight(marketSymbol)+symbolSeparator+TrimRight(companyName)
    else
      Cells[portfolioCol[0],RowCount-1]:=TrimRight(companyName)+symbolSeparator+TrimRight(marketSymbol);
    if company>0 then
      company:=(company-1) div nameLength
    else
      dec(company);
    Read(purchasePrice,4);
    Read(purchaseDate,2);
    Read(numShares,4);
	  Cells[portfolioCol[5],RowCount-1]:=priceToStr(purchasePrice);
    if purchaseDate>0 then
      Cells[portfolioCol[6],RowCount-1]:=IntToStr(purchaseDate and 31)+'-'
      +Copy(months,3*((purchaseDate shr 5) and 15)-2,3)+'-'
      +twoDigits((1950+(purchaseDate shr 9)) mod 100)
      //IntToStr(purchaseDate and 31)+'/'
      //+IntToStr(purchaseDate shr 5 and 15)+'/'+IntToStr(1950+(purchaseDate shr 9 and 127))
    else
      Cells[portfolioCol[6],RowCount-1]:=notAvailStr;
    Cells[portfolioCol[7],RowCount-1]:=IntToStr(numShares);
    if newStyle then
      begin
      Read(averages,3);
      if averages=0 then
        Cells[portfolioCol[10],RowCount-1]:=defaultStr
      else
        begin
        if (averages and 1)=1 then S:='S' else S:='E';
        Cells[portfolioCol[10],RowCount-1]:=S+'('+IntToStr((averages shr 1) and 511)
        +','+IntToStr((averages shr 10) and 511)+')';
        end;
      end
    else
      Cells[10,RowCount-1]:=defaultStr;
	  updatePortfolioColumns(RowCount-1,marketNo,company);
	  RowCount:=RowCount+1;
    end;
  portfolioChanged:=(not getOriginal);
  except
    Result:=false;
    portfolioChanged:=false;
    end;
  Free;
  editRow:=RowCount-1;
  RowHeights[0]:=2*Canvas.TextHeight('W')+1;
  for i:=0 to 10 do
    begin
    Cells[portfolioCol[i],0]:=columnLabels[i];
    Cells[i,RowCount-1]:='';
    end;
  btnRemove.Enabled:=(RowCount>2);
  if btnRemove.Enabled then updateRightControls(Cells[0,Row]);
  updatePortfolioTotal;
  btnSave.Enabled:=portfolioChanged or anyPricesChanged;
  //btnPortfolioUndo.Enabled:=portfolioChanged or FileExists(Copy(fileName,1,length(fileName)-4)+'.bak');
  DateTimePicker1.Hide;
  end;
end;
end;

function extractAvgSettings(const S:String;out avg1,avg2:Integer):Boolean;//result is true if avg type is simple
var
  commaPos:Integer;
begin
if S=defaultStr then
  begin
  avg1:=0;
  avg2:=0;
  Result:=false;
  end
else
  begin
  commaPos:=Pos(',',S);
  avg1:=StrToInt(Copy(S,3,commaPos-3));
  avg2:=StrToInt(Copy(S,commaPos+1,length(S)-commaPos-1));
  Result:=(S[1]='S');
  end;
end;

function wordToPrice(w:Word):Single;
begin
Result:=(w and 16383) + (w and 49152 shr 14)/4;
end;

function getSymbol(marketNo:Integer):String;
begin
Result:=Copy(marketSymbols,markets[marketNo]*nameLength+1,nameLength);
end;

function companyAndMarketNo(out market:Integer; companyName:String):Integer;
var
  name,symbol:String;
begin
if numMarkets>0 then
  begin
  if addSymbol then
    begin
    extractNameAndSymbol(companyName,name,symbol);
    market:=mktArrayIndex((Pos(symbol,marketSymbols)-1) div nameLength);
    Result:=(Pos(name,namesArray[market])-1) div nameLength;
    end
  else
    begin
    companyName:=padName(companyName);
    market:=numMarkets;
    Repeat
      dec(market);
      Result:=Pos(companyName,namesArray[market])-1;
    Until (Result>=0) or (market=0);
    if Result>=0 then Result:=Result div nameLength;
    end;
  end
else
  Result:=-1;
end;

function padName(const name:String):String;
begin
Result:=name+StringOfChar(' ',nameLength-length(name));
end;

{function companyNumber(const companyName:String):Integer;
begin
Result:=(Pos(companyName,namesArray[currMarket])-1) div nameLength;
end;}

procedure TForm1.lstMovingDownClick(Sender: TObject);
begin
if not hintDown(Sender) then
  begin
  lstMovingUp.ItemIndex:=-1;
  updateRightControls(lstMovingDown.Items[lstMovingDown.ItemIndex]);
  end;
end;

procedure TForm1.btnDeleteClick(Sender: TObject);
var
   i,shareNo,marketNo,noDays:Integer;
begin
if hintDown(Sender) then Exit;
Screen.Cursor:=crHourGlass;
i:=lstDelisted.Items.Count-1;
while i>=0 do
  begin
  if lstDelisted.Selected[i] then
    begin
    shareNo:=companyAndMarketNo(marketNo,lstDelisted.Items[i]);
    noDays:=length(yahooPrices[marketNo,shareNo]);
    if (noDays>=2) and (yahooPrices[marketNo,shareNo,noDays-1]<=0) and (yahooPrices[marketNo,shareNo,noDays-2]<=0) then
      begin
      removeFromPortfolio(lstDelisted.Items[i]);
      deleteShare(marketNo,shareNo);
      btnSave.Enabled:=true;
      pricesChanged[marketNo]:=true;
      dec(numDelisted[marketNo]);
      lstDelisted.Items.Delete(i);
      end;
    end;
  dec(i);
  end;
refreshLists;
btnDelete.Enabled:=False;
Screen.Cursor:=crDefault;
end;

procedure TForm1.removeFromPortfolio(const shareName:String);
var
	r:Integer;
begin
//remove every occurrence of a delisted share from grid
with grdPortfolio do
	for r:=RowCount-1 downto 1 do
    if (Cells[0,r]=shareName) and (Cells[portfolioCol[1],r]=notAvailStr) then
      removePortfolioShare(r);
updatePortfolioTotal;
end;

procedure deleteShare(market,shareNo:Integer);
var
	company:Integer;
begin
//noDays:=length(datesArray[market]);
for company:=shareNo+1 to numCompanies(market)-1 do
  begin
  SetLength(yahooPrices[market,company-1],length(yahooPrices[market,company]));
  Move(yahooPrices[market,company,0],yahooPrices[market,company-1,0],4*length(yahooPrices[market,company]));
  end;
Delete(namesArray[market],shareNo*nameLength+1,nameLength);
end;

procedure TForm1.savePortfolio;
var
  i,r,noRows,numShares,colOrder,averages,avg1,avg2: Integer;
  purchasePrice:Single;
  purchaseDate:Word;
  companyName,fileName,symbol:String;
  columnWidth:Byte;
  simpleAvgType:Boolean;
begin
fileName:=lstPortfolio.Items[editedIndex];
DeleteFile(fileName+'.bak');
RenameFile(fileName+'.dat',fileName+'.bak');
TFileStream.Create(fileName+'.dat',fmCreate).Free;
with TFileStream.Create(fileName+'.dat',fmOpenWrite or fmShareDenyNone) do
  begin
  try
  colOrder:=0;
  for i:=0 to 7 do
    colOrder:=colOrder or (portfolioCol[i] shl (i shl 2));
  Write(colOrder,4);
  colOrder:=1;
  Write(colOrder,1);//non-zero indicates newStyle
  dec(colOrder);//zero
  for i:=0 to 2 do
    colOrder:=colOrder or (portfolioCol[i+8] shl (i shl 2));
  Write(colOrder,2);
  with grdPortfolio do
    begin
    for i:=0 to 10 do
      begin
      columnWidth:=ColWidths[portfolioCol[i]];
      Write(columnWidth,1);
      end;
    noRows:=RowCount-1;
    for r:=1 to noRows-1 do
      begin
      //extract name and symbol from first portfolio column
      extractNameAndSymbol(Cells[portfolioCol[0],r],companyName,symbol);
      Write(Pointer(companyName)^,nameLength);
      Write(Pointer(symbol)^,nameLength);
      purchasePrice:=cellValue(Cells[portfolioCol[5],r]);
      purchaseDate:=dateStrToWord(Cells[portfolioCol[6],r]);
      numShares:=StrToInt(Cells[portfolioCol[7],r]);
      Write(purchasePrice,4);
      Write(purchaseDate,2);
      Write(numShares,4);
      if Cells[portfolioCol[10],r]=defaultStr then
        averages:=0
      else
        begin
        simpleAvgType:=extractAvgSettings(Cells[portfolioCol[10],r],avg1,avg2);
        averages:=Ord(simpleAvgType) or (avg1 shl 1) or (avg2 shl 10);
        end;
      Write(averages,3);
      end;
    end;
  portfolioChanged:=false;
  except end;
  Free;
  end;
end;

procedure extractNameAndSymbol(const S:String; out companyName,symbol:String);
var
  i:Integer;
begin
if prefix then
  begin
  i:=ScanC(S,' ',1);
  symbol:=padName(Copy(S,1,i-1));
  companyName:=padName(Copy(S,i+1,255));
  end
else
  begin
  i:=ScanB(S,' ',0);
  if (i>0) then
    begin
    symbol:=padName(Copy(S,i+1,255));
    companyName:=padName(Copy(S,1,i-1));
    end
  else
    begin
    symbol:=padName('');
    companyName:=padName(S);
    end;
  end;
end;

function dateStrToWord(const S:String):Word;
var
  year,month,day:Word;
begin
if S<>'-' then
  begin
  if S[3]='-' then
    begin
    day:=StrToInt(S[1]+S[2]);
    month:=(Pos(Copy(S,4,3),months)+2) div 3;
    year:=StrToInt(Copy(S,8,2));
    end
  else
    begin
    day:=StrToInt(S[1]);
    month:=(Pos(Copy(S,3,3),months)+2) div 3;
    year:=StrToInt(Copy(S,7,2));
    end;
  if year<50 then
    Result:=day or (month shl 5) or ((year+50) shl 9)
  else
    Result:=day or (month shl 5) or ((year-50) shl 9);
  end
else
  Result:=0;
end;

procedure TForm1.lstAllSharesClick(Sender: TObject);
var
  i:Integer;
begin
i:=lstAllShares.ItemIndex;
popupGraph.Autopopup:=(i>=0);
if hintDown(Sender) or (i<0) then Exit;
companyNo:=companyAndMarketNo(currMarket,lstAllShares.Items[i]);
selectedName:=lstAllShares.Items[i];
//selectedName:=getName(currMarket,companyNo);
PopUpMenu1.Autopopup:=true;
ShowFundamentals1.Visible:=(markets[currMarket]>0) and (markets[currMarket]<length(urlCountry));
//caption:=inttostr(companyNo);
btnAdd.Enabled:={(selectedName[1]<>'^') and }(lstPortfolio.ItemIndex>=0)
  and (lstPortfolio.ItemIndex<lstPortfolio.Items.Count-1);
AddToPortfolio.Enabled:=btnAdd.Enabled;
btnCombine.Enabled:=(lstDelisted.SelCount=1) and (selectedName[i]<>'^')
 and ((companyNo>numDelisted[currMarket]) or duplicateExists(currMarket,selectedName));
if sharesListed then
try
  refreshGraphSelection;
except
  Exit;
end;
populatePriceGrid;                                        
Case PageControl2.ActivePage.PageIndex of
  0: autoSelectAverageItem;
  1: autoSelectMoversItem;
//  2: lstDelisted.ItemIndex:=lstDelisted.Items.IndexOf(selectedName);
  3: autoSelectPortfolioItem;
  4: autoSelectHighLowItem;
  end;
end;

procedure TForm1.refreshGraphSelection;
begin
selectionEnd:=Max(0,length(yahooPrices[currMarket,companyNo])-1);
if lstRange.ItemIndex<7 then
  selectionStart:=Max(0,selectionEnd-rangeOptions[lstRange.ItemIndex])
else
  selectionStart:=0;
graphScroll.LargeChange:=Min(length(yahooPrices[currMarket,companyNo])-1,rangeOptions[lstRange.ItemIndex]);
with graphScroll do
{  if (Position=selectionEnd) and (Min=LargeChange) and (Max=selectionEnd) then
    refreshGraph
  else
    begin}
    SetParams(selectionEnd, LargeChange, selectionEnd);
    refreshGraph;
//    end;
showPrice.Enabled:=true;
end;

function duplicateExists(market:Integer; companyName:String):Boolean;
begin
Result:=IsFound(namesArray[market],padName(companyName),Pos(padName(companyName),namesArray[market])+nameLength);
end;

procedure TForm1.autoSelectHighLowItem;
begin
lstNearHigh.ItemIndex:=lstNearHigh.Items.IndexOf(selectedName);
lstNearLow.ItemIndex:=lstNearLow.Items.IndexOf(selectedName);
end;

procedure TForm1.autoSelectPortfolioItem;
var
   r: Integer;
   S,symbol:String;
   rect: TGridRect;
begin
if addSymbol then
  S:=selectedName
else
  begin
  symbol:=TrimRight(getSymbol(currMarket));
  if prefix then
    S:=symbol+symbolSeparator+selectedName
  else
    S:=selectedName+symbolSeparator+symbol;
  end;
With grdPortfolio do
  for r:=1 to RowCount-2 do
    if Cells[0,r]=S then
   	  begin
      rect.Left:=LeftCol;
      rect.Right:=VisibleColCount;
      rect.Top:=r;
      rect.Bottom:=r;
      Selection:=rect;
      Break;
      end;
end;

procedure TForm1.autoSelectAverageItem;
begin
lstMovingUp.ItemIndex:=lstMovingUp.Items.IndexOf(selectedName);
lstMovingDown.ItemIndex:=lstMovingDown.Items.IndexOf(selectedName);
end;

procedure TForm1.autoSelectMoversItem;
var
   i:Integer;
begin
i:=lstRisers.Items.Count;
Repeat
  dec(i)
Until (i<0) or (Copy(lstRisers.Items[i],1,length(lstRisers.Items[i])-8)=selectedName);
lstRisers.ItemIndex:=i;
i:=lstFallers.Items.Count;
Repeat
  dec(i)
Until (i<0) or (Copy(lstFallers.Items[i],1,length(lstFallers.Items[i])-8)=selectedName);
lstFallers.ItemIndex:=i;
end;

function cellValue(const editValue:String):Single;
var
   error1, error2: Integer;
begin
Val(editValue, Result, error1);
If error1>0 then
  begin
  If error1>1 then Val(Copy(editValue,1,error1-1),Result,error2);
  Case editValue[error1] of
    '¼': Result:=Result+0.25;
    '½': Result:=Result+0.5;
    '¾': Result:=Result+0.75;
    end;
  end;
end;

function numCompanies(marketNo:Integer):Integer;
begin
Result:=length(namesArray[marketNo]) div nameLength;
end;

procedure setSymbol(market:Integer);
begin
if addSymbol then
  begin
  mktSymbol:=TrimRight(getSymbol(market));
  if prefix then mktSymbol:=mktSymbol+symbolSeparator else mktSymbol:=symbolSeparator+mktSymbol;
  end
else
  mktSymbol:='';
end;

procedure TForm1.refreshLists;
var
  i,market: Integer;
begin
lstAllShares.Clear;
with lstAllShares.Items do
  for market:=0 to numMarkets-1 do if marketTicked(market) then
    begin
    setSymbol(market);
    if prefix then
      begin
      for i:=0 to numCompanies(market)-1 do if withinScope(market,i) then
        Add(mktSymbol+RTrim(getName(market,i),' '));
      end
    else
      begin
      for i:=0 to numCompanies(market)-1 do if withinScope(market,i) then
        Add(RTrim(getName(market,i),' ')+mktSymbol);
      end
    end;
i:=lstAllShares.Items.Indexof(selectedName);
lstAllShares.ItemIndex:=i;
if (i<0) then
  begin
  if (lstAllShares.Items.Count>0) then
    begin
    lstAllShares.ItemIndex:=0;
    companyNo:=companyAndMarketNo(currMarket,lstAllShares.Items[0]);
    sharesListed:=true;
    end
  else
    begin
    selectedName:='';
    sharesListed:=false;
    grdPrice.Rows[0].Clear;
    grdPrice.Rows[1].Clear;
    lblGraphTitle.Caption:='';
//    Graph.Canvas.Brush.Color:=clWhite;
//    Graph.Canvas.FillRect(Rect(0,0,Graph.Width,Graph.Height));
    GraphPaint(nil);
    end;
  grdPrice.Enabled:=sharesListed;
  end;
{btnSelectAll.Enabled:=false;
for market:=0 to numMarkets-1 do
  if lstMarkets.Checked[market] and (numDelisted[market]>0) then
    begin
    btnSelectAll.Enabled:=true;
    Break;
    end;}
//btnSelectAll.Enabled:=(numDelisted[market]>0);
lblAllShares.Caption:='Shares ('+ IntToStr(lstAllShares.Items.Count)+')';
lstDelisted.Clear;
for market:=0 to numMarkets-1 do if marketTicked(market) then
  begin
  setSymbol(market);
	for i:=1 to numDelisted[market]{-1} do
    lstDelisted.Items.Add(companyName(market,i));
  end;
btnSelectAll.Enabled:=lstDelisted.Items.Count>0;
lblDelisted.Caption:='Delisted? ('+ IntToStr(lstDelisted.Items.Count)+')';
updateRecentIssues;
populateUpDownLists;
updateMoversLists;
updatePortfolioGrid;
updateHighLowLists;
lstAllSharesClick(nil);
end;

procedure TForm1.updateRecentIssues;
var
	company,market,recentDays:Integer;
begin
lstRecentIssues.Clear;
recentDays:=spnRecentDays.Value;
for market:=0 to numMarkets-1 do
 if marketTicked(market) and (length(yahooPrices[market])>0) and (length(yahooPrices[market,0])>recentDays) then
  begin
  setSymbol(market);
  for company:=numDelisted[market]+1 to numCompanies(market)-1 do
    if withinScope(market,company) and (length(yahooPrices[market,company])<=recentDays) then
        lstRecentIssues.Items.Add(companyName(market,company));
  end;
end;

procedure TForm1.calculateGraph;
var
   day:Integer;
begin
//firstPoint:=selectionStart;
SetLength(graphData,selectionEnd+1);
//ensure enough data to calc high average later
if (simpleAvg) and (selectionStart>highAverage) then
  day:=selectionStart-highAverage+1
else
  day:=0;
while day<=selectionEnd do
  begin
  graphData[day,0]:=Abs(yahooPrices[currMarket,companyNo,day]);
  inc(day);
//  if graphData[day,0]=0 then graphData[day,0]:=graphData[day-1,0];
  end;
if (lowAverage>1) and (lowAverage<highAverage) then
  begin
  calcAverage(lowAverage,1);
  calcAverage(highAverage,2);
	end
else if (highAverage>1) then
  calcAverage(highAverage,1);
end;

function getFirstDay(average:Integer):Integer;
begin
if selectionStart>=average then Result:=selectionStart else Result:=average-1;
end;

procedure TForm1.calcAverage(spinSetting, arrayIndex: Integer);
var
   day, lastDay: Integer;
   factor,sum: Single;
begin
sum:=0;
if (selectionStart>spinSetting-1) then
  begin
  day:=selectionStart-spinSetting+1;
  lastDay:=selectionStart;//sum to first point
  end
else
  begin
  day:=0;
  lastDay:=spinSetting-1;//sum to spinSetting,as pre-firstPoint is zero
  end;
if lastDay>selectionEnd then Exit;//off graph boundary
//sum average to first point
Repeat
  sum:=sum+graphData[day,0];
  inc(day);
Until day>lastDay;
//both simple and exponential use simple moving average for first day
graphData[day-1,arrayIndex]:=sum/spinSetting;
if simpleAvg then
  //simple moving average
  While day<=selectionEnd do
    begin
    sum:=sum+graphData[day,0]-graphData[day-spinSetting,0];
    graphData[day,arrayIndex]:=sum/spinSetting;
    inc(day);
    end
else
  begin
  //exponential moving average
  factor:=2/(spinSetting+1);
  While day<=selectionEnd do
    begin
    graphData[day,arrayIndex]:=((graphData[day,0]-graphData[day-1,arrayIndex])*factor)+graphData[day-1,arrayIndex];
    inc(day);
    end;
  end;
end;

procedure TForm1.plotGraph(toScreen:Boolean);
var
	xFactor,yFactor,y:Single;
  day,y0,current,lastMark,nextMark,i,offset,noMonths,dayRange,xScale,graphEnd:Integer;
  graphHeight, graphWidth, arrayIndex, xAdjust, yAdjust, dateOffset, firstDay:Integer;
  leftMargin, rightMargin, topMargin, bottomMargin, divisionWidth: Integer;
  S:String;
  targetCanvas:TCanvas;
  showGrid,allData:Boolean;
begin
if selectedName='' then Exit;
dayRange:=selectionEnd-selectionStart+1;
noMonths:=dayRange div 21;
if toScreen then
  begin
  targetCanvas:=Graph.Canvas;
  graphWidth:=Graph.Width;
  graphHeight:=Graph.Height;
  leftMargin:=graphLeftMargin;
  rightMargin:=graphRightMargin;
  bottomMargin:=graphBottomMargin;
  topMargin:=graphTopMargin;
  xAdjust:=8;
  yAdjust:=6;
  end
else
  begin
  targetCanvas:=Printer.Canvas;
  graphWidth:=Printer.PageWidth;
  graphHeight:=Printer.PageHeight;
  leftMargin:=200;
  rightMargin:=100;
  topMargin:=250;
  bottomMargin:=150;
  xAdjust:=30;
  yAdjust:=40;
  end;
//xScale: 0=full text months, 1=abbrev text months, 2=years, 3=decades
if (targetCanvas.TextWidth('Dec')+9)*noMonths < graphWidth then
  xScale:=0
else if (targetCanvas.TextWidth('D')+9)*noMonths < graphWidth then
  xScale:=1
else if (targetCanvas.TextWidth('2000')+9)*noMonths div 12 < graphWidth then
  xScale:=2
else
  xScale:=3;

xFactor:=(graphWidth-leftMargin-rightMargin)/dayRange;

y0:=graphHeight-bottomMargin;
//determine lowest and highest marks required on y axis
yFactor:=(y0-topMargin)/(high-low);
lastMark:=leftMargin;

showGrid:=chkShowGrid.Checked;

With targetCanvas do begin
//draw x-axis
Pen.Color:=clBlack;
dateOffset:=length(datesArray[currMarket])-length(yahooPrices[currMarket,companyNo]);
firstDay:=dateOffset+selectionStart;
day:=firstDay;
graphEnd:=dateOffset+selectionEnd;
Repeat
  MoveTo(lastMark,graphHeight);
  LineTo(lastMark,y0);
  if chkShowGrid.Checked then
    begin
    Pen.Color:=clSilver;
    LineTo(lastMark,topMargin);
    Pen.Color:=clBlack;
    end;
  if xScale<=1 then
    begin
    current:=(datesArray[currMarket,day] shr 5) and 15;
    inc(day);
    while ((datesArray[currMarket,day] shr 5) and 15) = current do inc(day);
    if xScale=0 then
      S:=Copy(months,current*3 - 2,3)
    else
  		S:=months[current*3 - 2];
    end
  else
    begin
    current:=datesArray[currMarket,day] shr 9;
    inc(day);
    if xScale=2 then
      begin
      while (datesArray[currMarket,day] shr 9) = current do inc(day);
      S:=twoDigits(current mod 100);
      end
    else
      begin
      while (datesArray[currMarket,day] shr 9) div 10 = current div 10 do inc(day);
      S:=twoDigits((10*(current div 10)) mod 100)+'s';
      end;
    end;
  nextMark:=leftMargin+Round((day-firstDay)*xFactor);
  if nextMark>graphWidth-rightMargin then nextMark:=graphWidth-rightMargin;
  offset:=(nextMark+lastMark-TextWidth(S))div 2;
  if offset>lastMark then TextOut(offset,y0+2,S);
	lastMark:=nextMark;
Until day>=graphEnd;
//draw y-axis
i:=0;
While (high-low)/yIntervals[i]>13 do inc(i);
y:=yIntervals[i]*(Trunc(low/yIntervals[i]));
if y<low then y:=y+yIntervals[i];
Repeat
	offset:=y0-Round(yFactor*(y-low));
  MoveTo(9*leftMargin div 10,offset);
  LineTo(leftMargin,offset);
  if showGrid then
   	begin
    Pen.Color:=clSilver;
    MoveTo(leftMargin+1,offset);
   	LineTo(graphWidth-rightMargin,offset);
    Pen.Color:=clBlack;
    end;
  S:=FloatToStr(y);
  TextOut(leftMargin-TextWidth(S)-xAdjust,offset-yAdjust,S);
  y:=y+yIntervals[i];
Until y>high;
MoveTo(leftMargin,y0);
LineTo(leftMargin,topMargin-1);
MoveTo(leftMargin,y0);
LineTo(graphWidth-rightMargin,y0);
if (lowAverage>1) then
  arrayIndex:=3
else if (highAverage>1) then
 	arrayIndex:=2
else
  arrayIndex:=1;
inc(leftMargin);
Repeat
	dec(arrayIndex);
  if arrayIndex=0 then
    Pen.Color:=clBlack
  else if arrayIndex=1 then
    Pen.Color:=lblAvg1.Font.Color
  else
    Pen.Color:=lblAvg2.Font.Color;
  if (arrayIndex=2) or ((lowAverage=1) and (arrayIndex=1)) then day:=getFirstDay(highAverage)
  else if (arrayIndex=1) then day:=getFirstDay(lowAverage)
  else day:=selectionStart;
  if day<selectionEnd then
    begin
    MoveTo(Round((day-selectionStart)*xFactor)+leftMargin,y0-Round(yFactor*(graphData[day,arrayIndex]-low)));
    Repeat
  	  inc(day);
      //may be better to store (graphData - low) in graphData array
	    LineTo(Round((day-selectionStart)*xFactor)+leftMargin,y0-Round(yFactor*(graphData[day,arrayIndex]-low)));
    Until day=selectionEnd;
    end;
Until arrayIndex=0;
end;
end;

procedure TForm1.lstMovingUpClick(Sender: TObject);
begin
if not hintDown(Sender) then
  begin
  lstMovingDown.ItemIndex:=-1;
  updateRightControls(lstMovingUp.Items[lstMovingUp.ItemIndex]);
  end;
end;

procedure TForm1.lstDelistedClick(Sender: TObject);
begin
updateRightControls(lstDelisted.Items[lstDelisted.ItemIndex]);
btnDelete.Enabled:=(lstDelisted.SelCount>0);
//btnSelectAll.Enabled:=(lstDelisted.SelCount<>numDelisted[market]);
btnCombine.Enabled:=(lstDelisted.SelCount=1)
 and (lstDelisted.Items[lstDelisted.ItemIndex]<>lstAllShares.Items[lstAllShares.ItemIndex]);
//  or duplicateExists(market,lstDelisted.Items[lstDelisted.ItemIndex]));
end;

procedure TForm1.spnAverage1Change(Sender: TObject);
begin
if processSpinEdit(spnAverage1) and (sharesListed) then
  begin
  changeSpinValues(true);
  updatePortfolioGrid;
  populateUpDownLists;
  refreshGraph;
  end;
end;

procedure TForm1.changeSpinValues(averagesTab:Boolean);
var
  temp:Integer;
begin
if averagesTab then
  begin
  lowAverage:=spnAverage1.Value;
  highAverage:=spnAverage2.Value;
  simpleAvg:=(optSimple.Checked);
  end
else
  begin
  lowAverage:=spnTestAv1.Value;
  highAverage:=spnTestAv2.Value;
  btnGetResults.Enabled:=(lowAverage<>highAverage);
  simpleAvg:=optTestSimple.Checked;
  end;
If lowAverage>highAverage then
   begin
   temp:=lowAverage;
   lowAverage:=highAverage;
   highAverage:=temp;
   end;
if (lowAverage>1) and (lowAverage<highAverage) then
  begin
  lblAvg1.Show;
  lblAvg1.Caption:=IntToStr(lowAverage)+'-day';
  lblAvg2.Show;
  lblAvg2.Caption:=IntToStr(highAverage)+'-day';
  end
else if highAverage>1 then
  begin
  lblAvg2.Hide;
  lowAverage:=1;
  lblAvg1.Show;
  lblAvg1.Caption:=IntToStr(highAverage)+'-day';
  end
else
  begin
  highAverage:=1;
  lblAvg1.Hide;
  lblAvg2.Hide;
  end;
//refreshGraph;
end;

function TForm1.processSpinEdit(var spinControl:TSpinEdit):Boolean;
var
	newSpin,error: Integer;
begin
Val(spinControl.Text, newSpin, error);
Result:=(error=0) and (newSpin>=spinControl.MinValue) and (newSpin<=spinControl.MaxValue);
end;

procedure rangeCheck(var varToCheck:Integer;lowerLimit,upperLimit:Integer);
begin
if varToCheck<lowerLimit then
	varToCheck:=lowerLimit
else if varToCheck>upperLimit then
	varToCheck:=upperLimit;
end;

procedure TForm1.populateUpDownLists;
begin
lstMovingUp.Items.Clear;
lstMovingDown.Items.Clear;
if (lowAverage<highAverage) then
  begin
  if optSimple.Checked then
    simpleAveragePopulate
  else
    exponentialAveragePopulate;
  end;
lblMovingDown.Caption:='Below ('+IntToStr(lstMovingDown.Items.Count)+')';
lblMovingUp.Caption:='Above ('+IntToStr(lstMovingUp.Items.Count)+')';
autoSelectAverageItem;
//updatePortfolioGrid;
end;

function TForm1.marketTicked(marketNo:Integer):Boolean;
begin
if markets[marketNo]>=0 then
  Result:=lstMarkets.Checked[markets[marketNo]]
else
  Result:=false;
end;

procedure TForm1.simpleAveragePopulate;
var
   company, day, spin1Limit, spin2Limit,market,noDays, crossingDaysLimit,crossingDays: Integer;
   average1, average2: Single;
   parallelTest, below: Boolean;
   recentCrossChecked,risingParallel,fallingParallel: Boolean;
begin
recentCrossChecked:=chkRecentCrossing.Checked;
risingParallel:=chkRising.Checked;
fallingParallel:=chkFalling.Checked;
for market:=0 to numMarkets-1 do if marketTicked(market) then
  begin
  setSymbol(market);
  for company:=numDelisted[market]+1 to numCompanies(market)-1 do
    begin
    noDays:=length(yahooPrices[market,company]);
    if recentCrossChecked then
      begin
      crossingDays:=spnCrossing.Value;
      rangeCheck(crossingDays,1,noDays);
      crossingDaysLimit:=noDays-crossingDays;//crossingDays is zero if not recentCrossChecked
      end
    else
      crossingDaysLimit:=noDays;
    spin1Limit:=noDays-lowAverage;
    spin2Limit:=noDays-highAverage;
    if (crossingDaysLimit>=0) and (spin2Limit>=0) {and (spin1Limit>=0) }and withinScope(market,company)
    { and (price(market,company,spin2Limit)<>0)} then
      begin
      average1:=0;
      day:=noDays;
      Repeat
        dec(day);
        average1:=average1+Abs(yahooPrices[market,company,day]);
      Until day=spin1Limit;
      average2:=average1;
      Repeat
        dec(day);
        average2:=average2+Abs(yahooPrices[market,company,day]);
      Until day=spin2Limit;
      below:=(average1/lowAverage<=average2/highAverage);
      //determine if average2 and average1 are parallel
      if ((not below) and risingParallel) or (below and fallingParallel) then
        begin
        day:=noDays;
        Repeat
          dec(day);
        Until Abs(yahooPrices[market,company,day-lowAverage])<>Abs(yahooPrices[market,company,day]);
        parallelTest:=(below=(Abs(yahooPrices[market,company,day])<=Abs(yahooPrices[market,company,day-lowAverage])));
        if parallelTest then
          begin
          day:=noDays;
          Repeat dec(day) Until Abs(yahooPrices[market,company,day-highAverage])<>Abs(yahooPrices[market,company,day]);
          parallelTest:=(below=(Abs(yahooPrices[market,company,day])<=Abs(yahooPrices[market,company,day-highAverage])));
          end;
        end
      else
        parallelTest:=true;
      if parallelTest then
        begin
        day:=noDays;
        While (day>crossingDaysLimit) and (below xor (average1/lowAverage>average2/highAverage)) do
          begin
          dec(day);
          average1:=average1+Abs(yahooPrices[market,company,day-lowAverage])-Abs(yahooPrices[market,company,day]);
          average2:=average2+Abs(yahooPrices[market,company,day-highAverage])-Abs(yahooPrices[market,company,day]);
          end;
        if (not recentCrossChecked) or (below xor (average1/lowAverage<=average2/highAverage)) then
          begin
          If (below) then
            lstMovingDown.Items.Add(companyName(market,company))
          else
            lstMovingUp.Items.Add(companyName(market,company));
          end;
        end;
      end;
    end;
  end;
end;

procedure TForm1.exponentialAveragePopulate;
var
   company, day, market, noDays,crossingDaysLimit,crossingDays: Integer;
   average1, average2: Single;
   oldAverage, averageStore: Single;
   parallelTest, below: Boolean;
   recentCrossChecked,risingParallel,fallingParallel: Boolean;
begin
recentCrossChecked:=chkRecentCrossing.Checked;
risingParallel:=chkRising.Checked;
fallingParallel:=chkFalling.Checked;
factor1:=2/(lowAverage+1);
factor2:=2/(highAverage+1);
factor1min1:=1-factor1;
factor2min1:=1-factor2;
for market:=0 to numMarkets-1 do if marketTicked(market) then
  begin
  setSymbol(market);
  for company:=numDelisted[market]+1 to numCompanies(market)-1 do
    begin
    noDays:=length(yahooPrices[market,company]);
    if recentCrossChecked then
      begin
      crossingDays:=spnCrossing.Value;
      rangeCheck(crossingDays,1,noDays);
      crossingDaysLimit:=noDays-crossingDays;//crossingDays is zero if not recentCrossChecked
      end
    else
      crossingDaysLimit:=noDays;
    if withinScope(market,company) and (noDays>=highAverage) and (yahooPrices[market,company,noDays-highAverage]<>0)
        and calcExpAverage(market,company,average1,average2) then
      begin
      //sum initial simple moving average
      below:=(average1<=average2);
      //determine if average2 and average1 are parallel
      if ((not below) and risingParallel) or (below and fallingParallel) then
        begin
        day:=noDays-1;
        if factor1min1>0 then
          begin
          averageStore:=average1;
          Repeat
            oldAverage:=average1;
            average1:=(average1-factor1*Abs(yahooPrices[market,company,day]))/factor1min1;
            dec(day);
          Until average1<>oldAverage;
          parallelTest:=(below=(oldAverage<=average1));
          average1:=averageStore;
          end
        else //when period is 1, EMA=price
          begin
          While Abs(yahooPrices[market,company,day])=Abs(yahooPrices[market,company,day-1]) do dec(day);
          parallelTest:=(below=(Abs(yahooPrices[market,company,day])<=Abs(yahooPrices[market,company,day-1])));
          end;
        if parallelTest then
          begin
          day:=noDays-1;
          averageStore:=average2;
          Repeat
            oldAverage:=average2;
            average2:=(average2-factor2*Abs(yahooPrices[market,company,day]))/factor2min1;
            dec(day);
          Until average2<>oldAverage;
          parallelTest:=(below=(oldAverage<=average2));
          average2:=averageStore;
          end;
        end
      else
        parallelTest:=true;
      if parallelTest then
        begin
        day:=noDays-1;
        //work backwards from current moving averages to find crossing point
        if factor1min1>0 then
          Repeat
            average1:=(average1-factor1*Abs(yahooPrices[market,company,day]))/factor1min1;
            average2:=(average2-factor2*Abs(yahooPrices[market,company,day]))/factor2min1;
            dec(day);
          Until (day<crossingDaysLimit) or (below xor (average1<=average2))
        else
          //note: average1 is already price(pricesArray[company,day]);
          Repeat
            average2:=(average2-factor2*average1)/factor2min1;
            dec(day);
            average1:=Abs(yahooPrices[market,company,day]);
          Until (day<crossingDaysLimit) or (below xor (average1<=average2));
        if (not recentCrossChecked) or (below xor (average1<=average2)) then
          begin
          If (below) then
            lstMovingDown.Items.Add(companyName(market,company))
          else
            lstMovingUp.Items.Add(companyName(market,company));
          end;
        end;
      end;
    end;
  end;
end;

function calcExpAverage(market,company:Integer; out average1, average2: Single):Boolean;
var
  day,firstData,limit,noDays:Integer;
begin
day:=0;
average1:=0;
average2:=0;
//while price(market,company,day)=0 do inc(day);
noDays:=length(yahooPrices[market,company]);
Result:=(day+highAverage<noDays);
if not Result then exit;
firstData:=day;
if lowAverage>1 then
  begin
  limit:=firstData+lowAverage;
  Repeat
    average1:=average1+Abs(yahooPrices[market,company,day]);
    inc(day);
  Until day=limit;
  limit:=firstData+highAverage;
  average2:=average1;
  Repeat
    average2:=average2+Abs(yahooPrices[market,company,day]);
    inc(day);
  Until day=limit;
  average1:=average1/lowAverage;
  average2:=average2/highAverage;
  //now sum both exp moving averages up to present day
  day:=firstData+lowAverage;
  limit:=firstData+highAverage;
  Repeat
    average1:=((Abs(yahooPrices[market,company,day])-average1)*factor1)+average1;
    inc(day);
  Until day=limit;
  Repeat
    average1:=((Abs(yahooPrices[market,company,day])-average1)*factor1)+average1;
    average2:=((Abs(yahooPrices[market,company,day])-average2)*factor2)+average2;
    inc(day);
  Until day=noDays;
  end
else
  begin
  average1:=Abs(yahooPrices[market,company,noDays-1]);
  limit:=firstData+highAverage;
  day:=0;
  average2:=0;
  Repeat
    average2:=average2+Abs(yahooPrices[market,company,day]);
    inc(day);
  Until day=limit;
  average2:=average2/highAverage;
  While day<noDays do
    begin
    average2:=((Abs(yahooPrices[market,company,day])-average2)*factor2)+average2;
    inc(day);
    end;
  end;
end;

function withinScope(market, company:Integer):Boolean;
begin
if (scopeSetting and 2)=0 then
  Result:=true
else if (scopeSetting and 1)=0 then
  Result:=(Abs(yahooPrices[market,company,length(yahooPrices[market,company])-1])>threshold)
else
  Result:=(Abs(yahooPrices[market,company,length(yahooPrices[market,company])-1])<threshold);
end;

procedure saveMarket(const fileName:String; marketNo: Integer);
var
  r:Integer;
  noCompanies,noDays:Word;
begin
DeleteFile(fileName+'.bak');
RenameFile(fileName+'.mkt',fileName+'.bak');
TFileStream.Create(fileName+'.mkt',fmCreate).Free;
With TFileStream.Create(fileName+'.mkt',fmOpenWrite or fmShareDenyNone) do
  begin
  try
  noCompanies:=numCompanies(marketNo);
  Write(noCompanies,2);
  //write number of delisted shares
  Write(numDelisted[marketNo],2);
  noDays:=length(datesArray[marketNo]);
  Write(noDays,2);
  Write(Pointer(namesArray[marketNo])^,noCompanies*nameLength);
  Write(Pointer(datesArray[marketNo])^,noDays*2);
  for r:=0 to noCompanies-1 do
    begin
    noDays:=length(yahooPrices[marketNo,r]);
    Write(noDays,2);
    Write(Pointer(yahooPrices[marketNo,r])^,4*noDays);
    end;
  pricesChanged[marketNo]:=false;
  except end;
  Free;
  end;
end;

procedure TForm1.btnSaveClick(Sender: TObject);
var
  m:Integer;
  fileName:String;
begin
if hintDown(Sender) then Exit;
for m:=0 to numMarkets-1 do if pricesChanged[m] then
  begin
  fileName:=TrimRight(getSymbol(m));
  saveMarket(fileName,m);
  end;
if portfolioChanged then savePortfolio;
btnSave.Enabled:=false;
end;

procedure TForm1.btnCombineClick(Sender: TObject);
var
	i,oldCompanyNo,oldMarketNo:Integer;
begin
if hintDown(Sender) or (numMarkets=0) then Exit;
oldCompanyNo:=companyAndMarketNo(oldMarketNo,lstDelisted.Items[lstDelisted.ItemIndex]);
//companyNo is company selected in AllShares list
i:=length(yahooPrices[currMarket,companyNo]);
Repeat
dec(i);
if yahooPrices[currMarket,companyNo,i]=0 then
  begin
  yahooPrices[currMarket,companyNo,i]:=yahooPrices[oldMarketNo,oldCompanyNo,i];
  if (yahooPrices[currMarket,companyNo,i]=0) then
    yahooPrices[currMarket,companyNo,i]:=yahooPrices[currMarket,companyNo,i+1];
  end;
Until i=0;
populatePriceGrid;
spnAverage1Change(Sender);
btnCombine.Enabled:=False;
btnSave.Enabled:=true;
pricesChanged[currMarket]:=true;
end;

procedure TForm1.populatePriceGrid;
var
   day, lastDay, noDays, dateOffset:Integer;
   p:Single;
begin
With grdPrice do begin
noDays:=length(yahooPrices[currMarket,companyNo]);
ColCount:=noDays;
Refresh;
day:=LeftCol;
lastDay:=day+VisibleColCount;
{If selectedName=lseAvgStr then
  begin
  dayChange:=Round(100*allShare[numDays-1]/allShare[numDays-2])-100;
  If dayChange<=0 then changeSymbol:='' else changeSymbol:='+';
  //display latest price + %change on yesterday
  lblGraphTitle.Caption:=lseAvgStr+' ('+changeSymbol+IntToStr(dayChange)+'%)';
  while day<=lastDay do
    begin
    Cells[day,0]:=dayToStr(0,day);
    Cells[day,1]:=priceWithFraction(allShare[day]);
    inc(day);
    end;
  end
else
begin}
updateGraphTitle(noDays);
dateOffset:=length(datesArray[currMarket])-noDays;
if companyNo>=0 then while day<=lastDay do
  begin
  Cells[day,0]:=dayToStr(dateOffset+day);
  p:=yahooPrices[currMarket,companyNo,day];
  if p>0 then
    Cells[day,1]:=priceToStr(p)
  else
    Cells[day,1]:='-';
  inc(day);
  end;
Col:=noDays-1;
end;
end;

procedure updateGraphTitle(noDays:Integer);
var
  title: String;
  dayChange: Integer;
  changeSymbol:String[1];
begin
title:=StringReplace(selectedName,'&','&&',[rfReplaceAll]);
//display share name + latest price + %change on yesterday + sector
if ((companyNo=0) or (companyNo>numDelisted[currMarket])) and (yahooPrices[currMarket,companyNo,noDays-1]<>0) then
	begin
  If (noDays>1) and (yahooPrices[currMarket,companyNo,noDays-2]<>0) then
   	dayChange:=Round(100*Abs(yahooPrices[currMarket,companyNo,noDays-1])/Abs(yahooPrices[currMarket,companyNo,noDays-2]))-100
	else
   	dayChange:=0;
  If dayChange<=0 then changeSymbol:='' else changeSymbol:='+';
  title:=title+'  '+changeSymbol+IntToStr(dayChange)+'% @ '
  +priceToStr(Abs(yahooPrices[currMarket,companyNo,noDays-1]));
  if (lowAverage<highAverage) and (noDays>=highAverage) then
    begin
    if belowAverage(noDays,currMarket,companyNo,simpleAvg,lowAverage,highAverage) then
      title:=title+'  (-)'
    else
      title:=title+'  (+)';
	  end;
  end
else
	title:=title+' (Delisted?)';
Form1.lblGraphTitle.Caption:=title;
end;

function dayToStr(day:Integer):String;
begin
Result:=IntToStr(datesArray[currMarket,day] and 31)+'-'
+Copy(months,3*((datesArray[currMarket,day] shr 5) and 15)-2,3)+'-'
+twoDigits((datesArray[currMarket,day] shr 9) mod 100);
//Result:=IntToStr(datesArray[market,day] and 31)+'/'+IntToStr((datesArray[market,day] shr 5) and 15)
//    +'/'+twoDigits((datesArray[market,day] shr 9) mod 100)
end;

function dayToStr(market,day:Integer):String;
begin
Result:=IntToStr(datesArray[market,day] and 31)+'-'
+Copy(months,3*((datesArray[market,day] shr 5) and 15)-2,3)+'-'
+twoDigits((datesArray[market,day] shr 9) mod 100);
end;

{function priceWithFraction(p: Integer):String;overload;
begin
If p>16383 then
  begin
  If (p and 16383)>0 then
     Result:=IntToStr(p and 16383) + fractions[p and 49152 shr 14]
  else
     Result:=fractions[p and 49152 shr 14];
  end
else if p>0 then
  Result:=IntToStr(p)
else
  Result:='-';
end;

function priceWithFraction(p: Single):String;overload;
var
	f:Integer;
begin
f:=Round(4*Frac((Abs(p))));
if (f<4) then
   begin
	 if Abs(p)>=1 then Result:=IntToStr(Trunc(p)) else Result:='';
   if f>0 then Result:=Result+fractions[f];
   if Result='' then Result:='0';
   end
else
	Result:=IntToStr(Trunc(p+1));
end; }

procedure TForm1.btnExitClick(Sender: TObject);
begin
if not hintDown(Sender) then Form1.Close;
end;

procedure TForm1.btnHelpClick(Sender: TObject);
begin
helpWindow:=HtmlHelp(GetDesktopWindow, '..\AutoShare.chm', HH_DISPLAY_TOPIC, 0);
//Application.HelpCommand(HELP_CONTENTS,0);
end;

procedure TForm1.btnUndoClick(Sender: TObject);
var
  m:Integer;
  problems,mktName:String;
begin
if hintDown(Sender) then Exit;
problems:='';
for m:=0 to numMarkets-1 do if length(namesArray[m])>0 then
  begin
  mktName:=TrimRight(getSymbol(m));
  if not getPricesFile(mktName,m,pricesChanged[m]) then
    begin
    getPricesFile(mktName,m,not pricesChanged[m]);
    problems:=problems+mktName+', ';
    end;
  end;
if problems>'' then
  errorMessage('The alternate price files for the following indices are damaged or cannot be found: '
  +Copy(problems,1,length(problems)-2));
refreshLists;
end;

procedure TForm1.FormResize(Sender: TObject);
var
  pageControlHeight,sheetHeight:Integer;
begin
pageControlHeight:=Form1.Height-86;
PageControl2.Height:=pageControlHeight;
sheetHeight:=highLowSheet.Height;
if PageControl2.Width=249 then
  begin
  Graph.Left:=384;
  Graph.Width:=Form1.Width-399;
  lblRange.Left:=220+(Graph.Width div 2);
  grdPortfolio.Height:=sheetHeight-94;
  end
else
  begin
  Graph.Left:=480;
  Graph.Width:=Form1.Width-495;
  lblRange.Left:=320+(Graph.Width div 2);
  grdPortfolio.Height:=sheetHeight-94;
  end;
grdPrice.Left:=Graph.Left;
lblGraphTitle.Left:=Graph.Left;
grdPrice.Width:=Graph.Width;
lblGraphTitle.Width:=Graph.Width;
lstRange.Left:=lblRange.Left+48;
chkShowGrid.Left:=lblRange.Left+156;
lblAvg1.Left:=lblRange.Left+220;
lblAvg2.Left:=lblAvg1.Left+58;
lblAvg1.Top:=pageControlHeight+31;
lblAvg2.Top:=lblAvg1.Top;
chkShowGrid.Top:=lblAvg1.Top-1;
lblRange.Top:=lblAvg1.Top;
lstRange.Top:=lblAvg1.Top-4;
lstAllShares.Height:=pageControlHeight-64;
lstMovingUp.Height:=sheetHeight-126;
lstMovingDown.Height:=lstMovingUp.Height;
chkRising.Top:=lstMovingUp.Height+107;
chkFalling.Top:=chkRising.Top;
lstRisers.Height:=(sheetHeight-114) div 2;
lstFallers.Height:=lstRisers.Height;
lblFallers.Top:=lstRisers.Height+95;
lstFallers.Top:=lblFallers.Top+16;
lstDelisted.Height:=sheetHeight - 118;
btnDelete.Top:=sheetHeight - 29;
btnSelectAll.Top:=sheetHeight - 58;
lstRecentIssues.Height:=sheetHeight - 60;
lblTotal.Top:=grdPortfolio.Height+40;
btnRemove.Top:=lblTotal.Top+25;
btnAdd.Top:=btnRemove.Top;
btnPortfolioUndo.Top:=btnRemove.Top;
lstNearHigh.Height:=sheetHeight-60;
lstNearLow.Height:=lstNearHigh.Height;
lstMarkets.Height:=pageControlHeight-198;
if graphScroll.Visible then
  Graph.Height:=pageControlHeight-94
else
  Graph.Height:=pageControlHeight-76;
graphScroll.Top:=pageControlHeight+5;
graphScroll.Width:=Graph.Width-1;
graphScroll.Left:=Graph.Left+1;
btnImport.Top:=pageControlHeight-168;
btnExport.Top:=btnImport.Top;
btnClear.Top:=btnImport.Top;
chkIndexOnly.Top:=pageControlHeight-136;
chkUnadjusted.Top:=pageControlHeight-118;
chkScope.Top:=pageControlHeight-68;
lstScope.Top:=chkScope.Top-2;
spnThreshold.Top:=lstScope.Top;
chkSymbol.Top:=pageControlHeight-94;
lstSymbol.Top:=chkSymbol.Top-2;
lblSymbol.Top:=chkSymbol.Top+2;
if Image1.Visible then
  begin
  Memo1.Height:=backTestSheet.Height-419;
  Image1.Top:=198+Memo1.Height;
  end
else
  Memo1.Height:=backTestSheet.Height-195;
if selectedName>'' then populatePriceGrid;
end;

procedure TForm1.btnSelectAllClick(Sender: TObject);
var
   i: Integer;
begin
if hintDown(Sender) {or (numDelisted[market]=0)} then Exit;
for i:=0 to lstDelisted.Items.Count-1 do lstDelisted.Selected[i]:=True;
btnSelectAll.Enabled:=False;
btnDelete.Enabled:=True;
end;

function companyName(market,company:Integer):String;
begin
if addSymbol then
  begin
  if prefix then
    Result:=mktSymbol+RTrim(getName(market,company),' ')
  else
    Result:=RTrim(getName(market,company),' ')+mktSymbol;
  end
else
  Result:=RTrim(getName(market,company),' ');
end;

function getName(marketNo,company:Integer):String;
begin
Result:=Copy(namesArray[marketNo],company*nameLength+1,nameLength);
end;

function dateToWord(day,month,year:Word):Word;
begin
Result:=day + (month shl 5) + ((year-1900{) mod 100}) shl 9);
end;

procedure TForm1.updateMoversLists;
var
 day, month, year:Word;
 startDate, lastIndexDate, endDate, lastDate, lastShareDate, dateOffset, change, company, specStartDate, specEndDate:Integer;
 startPrice, endPrice:Single;
 risersCount, fallersCount, market, i, period:Integer;
 changeStr1,changeStr2:String;//[16];
 risersList,fallersList:TStringList;
begin
risersList:=TStringList.Create;
fallersList:=TStringList.Create;
lstRisers.Clear;
lstFallers.Clear;
period:=comboPeriod.ItemIndex;
if period<8 then
  begin
  DecodeDate(pickerFrom.Date,year,month,day);
  specStartDate:=dateToWord(day,month,year);
  DecodeDate(pickerTo.Date,year,month,day);
  specEndDate:=dateToWord(day,month,year);
  end;
for market:=0 to numMarkets-1 do if marketTicked(market) then
  begin
  setSymbol(market);
  lastIndexDate:=length(datesArray[market])-1;
  if (lastIndexDate>0) and ((period=8) or (datesArray[market,0]<=specEndDate))
   and ((period>0) or (datesArray[market,lastIndexDate]>=specStartDate)) then
    begin
    lastDate:=lastIndexDate;
    if period<8 then
      begin
      while (lastDate>0) and (datesArray[market,lastDate]>specEndDate) do dec(lastDate);
      startDate:=lastDate-1;
      while (startDate>0) and (datesArray[market,startDate]>specStartDate) do dec(startDate);
      end;
    for company:=numDelisted[market]+1 to numCompanies(market)-1 do
      begin
      if withinScope(market,company) then
		    begin
        if period=8 then
          begin
   	      startPrice:=Abs(yahooPrices[market,company,0]);
          endDate:=length(yahooPrices[market,company])-1;
          while (endDate>0) and (yahooPrices[market,company,endDate]=0) do dec(endDate);
          if endDate>=0 then
            endPrice:=Abs(yahooPrices[market,company,endDate])
          else
            endPrice:=0;
          end
        else
          begin
          lastShareDate:=length(yahooPrices[market,company])-1;
          dateOffset:=lastShareDate-lastIndexDate;
          endDate:=lastDate+dateOffset;
          while (endDate>0) and (yahooPrices[market,company,endDate]=0) do dec(endDate);
          if (startDate+dateOffset>=0) and (endDate>=0) then
			      begin
   		      startPrice:=Abs(yahooPrices[market,company,startDate+dateOffset]);
   		      endPrice:=Abs(yahooPrices[market,company,endDate]);
            end
          else
            startPrice:=0;
          end;
   	    If (startPrice>0) and (endPrice>0) then
    	    begin
    	    change:=Round(100* endPrice/startPrice)-100;
          changeStr1:=IntToStr(Abs(change));
    	    If (change>0) then
       	    risersList.Add(StringOfChar(' ',6-length(changeStr1))+changeStr1+'%'+companyName(market,company))
    	    else if (change<0) then
       	    fallersList.Add(StringOfChar(' ',6-length(changeStr1))+changeStr1+'%'+companyName(market,company));
    	    end;
        end;
      end;
    end;
  end;
risersList.Sort;
fallersList.Sort;
risersCount:=risersList.Count;
lblRisers.Caption:='Risers ('+IntToStr(risersCount)+')';
If risersCount>0 then
   begin
   Dec(risersCount);
   for i:=0 to risersCount div 2 do
     begin
     changeStr1:=risersList.Strings[i];
     changeStr2:=risersList[risersCount-i];
     risersList.Strings[i]:=reverseMoverStr(changeStr2);
     risersList.Strings[risersCount-i]:=reverseMoverStr(changeStr1);
     end;
   lstRisers.Items.AddStrings(risersList);
   end;
fallersCount:=fallersList.Count;
lblFallers.Caption:='Fallers ('+IntToStr(fallersCount)+')';
If fallersCount>0 then
   begin
   Dec(fallersCount);
   for i:=0 to fallersCount div 2 do
     begin
     changeStr1:=fallersList.Strings[i];
     changeStr2:=fallersList.Strings[fallersCount-i];
     fallersList.Strings[i]:=reverseMoverStr(changeStr2);
     fallersList.Strings[fallersCount-i]:=reverseMoverStr(changeStr1);
     end;
   lstFallers.Items.AddStrings(fallersList);
   end;
autoSelectMoversItem;
end;

function reverseMoverStr(const moverStr:String):String;
var
  i: Integer;
begin
i:=2;
while moverStr[i]<>'%' do inc(i);
Result:=Copy(moverStr,i+1,255)+chr(9)+Copy(moverStr,1,i);
end;

procedure TForm1.lstRisersClick(Sender: TObject);
var
  name:String;
  i:Integer;
begin
if hintDown(Sender) then Exit;
lstFallers.ItemIndex:=-1;
name:=lstRisers.Items[lstRisers.ItemIndex];
i:=length(name)-2;
while name[i]<>chr(9) do dec(i);
updateRightControls(Copy(name,1,i-1));
end;

{function stripSymbol(const S:String):String;
var
  i:Integer;
begin
if addSymbol then
  begin
  if prefix then
    begin
    i:=1;
    Repeat inc(i) Until S[i]='.';
    Result:=Copy(S,i+1,255);
    end
  else
    begin
    i:=length(S);
    Repeat dec(i) Until S[i]='.';
    Result:=Copy(S,1,i-1);
    end;
  end
else
  Result:=S;
end;}

procedure TForm1.updateRightControls(const name:String);
begin
if lstAllShares.Items.IndexOf(name)>=0 then
  begin
  lstAllShares.ItemIndex:=lstAllShares.Items.IndexOf(name);
  companyNo:=companyAndMarketNo(currMarket,name);
  selectedName:=name;
  refreshGraphSelection;
  populatePriceGrid;
  PopUpMenu1.Autopopup:=true;
  ShowFundamentals1.Visible:=(lstAllShares.ItemIndex>=0) and (markets[currMarket]>0) and (markets[currMarket]<length(urlCountry));
  end
else
  PopUpMenu1.Autopopup:=false;
end;

procedure TForm1.refreshGraph;
var
  day, arrayIndex: Integer;
begin
updateGraphTitle(length(yahooPrices[currMarket,companyNo]));
calculateGraph;
//determine low and high limits of y-axis
low:=graphData[selectionStart,0];
high:=low;
if (lowAverage>1) then
  arrayIndex:=3
else if (highAverage>1) then
 	arrayIndex:=2
else
  arrayIndex:=1;
Repeat
	dec(arrayIndex);
  if (arrayIndex=2) or ((lowAverage=1) and (arrayIndex=1)) then day:=getFirstDay(highAverage)
  else if (arrayIndex=1) then day:=getFirstDay(lowAverage)
  else day:=selectionStart;
  While day<=selectionEnd do
    begin
    low:=Min(low,graphData[day,arrayIndex]);
    high:=Max(high,graphData[day,arrayIndex]);
    inc(day);
    end;
Until arrayIndex=0;
//round high and low to nearest quarter
high:=(1+Trunc(4*high))/4;//round up
if zeroY then
  low:=0
else
  begin
  if (high>low) or (frac(high)>0) or (high<10) then
	  low:=(Round(4*low)-1)/4 //place 1/4 unit above x-axis
  else
	  low:=high-10;
  end;
Graph.Invalidate;//plot graph
end;

procedure TForm1.lstFallersClick(Sender: TObject);
var
  name:String;
  i:Integer;
begin
if hintDown(Sender) then Exit;
lstRisers.ItemIndex:=-1;
name:=lstFallers.Items[lstFallers.ItemIndex];
i:=length(name)-2;
while name[i]<>chr(9) do dec(i);
updateRightControls(Copy(name,1,i-1));
end;

procedure TForm1.comboPeriodChange(Sender: TObject);
begin
Case comboPeriod.ItemIndex of
     1: pickerFrom.Date:=pickerTo.Date-1;
     2: pickerFrom.Date:=pickerTo.Date-7;
     3: pickerFrom.Date:=IncMonth(pickerTo.Date,-1);
     4: pickerFrom.Date:=IncMonth(pickerTo.Date,-3);
     5: pickerFrom.Date:=IncMonth(pickerTo.Date,-12);
     6: pickerFrom.Date:=IncMonth(pickerTo.Date,-24);
     7: pickerFrom.Date:=IncMonth(pickerTo.Date,-60);
     end;
updateMoversLists;
end;

procedure TForm1.pickerToCloseUp(Sender: TObject);
begin
pickerTo.MaxDate:=Date;
comboPeriodChange(Sender);
end;

procedure TForm1.pickerFromCloseUp(Sender: TObject);
begin
If pickerFrom.Date>=PickerTo.Date then pickerFrom.Date:=pickerTo.Date-1;
comboPeriod.ItemIndex:=0;
updateMoversLists;
end;

procedure TForm1.averagesSheetShow(Sender: TObject);
begin
if sharesListed then
  begin
  autoSelectAverageItem;
  refreshGraph;
  end;
end;

procedure TForm1.moversSheetShow(Sender: TObject);
begin
autoSelectMoversItem;
end;

procedure TForm1.lstAllSharesMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
if (not btnHint.Down) and (Button=mbLeft) and (portfolioSheet.Showing) and (lstAllShares.ItemIndex>0) then
  lstAllShares.BeginDrag(False);
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
   response:Word;
   i:Integer;
   changes:String;
begin
if btnSave.Enabled then
  begin
  if portfolioChanged then changes:=lstPortfolio.Text+', ' else changes:='';
  for i:=0 to numMarkets-1 do if pricesChanged[i] then
    changes:=changes+TrimRight(getSymbol(i))+', ';
  if changes>'' then
    begin
    response:=MessageDlg('Do you want to save changes to the following: '
    +Copy(changes,1,length(changes)-2)+'?',mtConfirmation, [mbYes, mbNo, mbCancel], 0);
    if response = mrCancel then
      CanClose:=False
    else if response = mrYes then
      btnSaveClick(Sender);
    end;
  end
else
  CanClose:=btnExit.Enabled;
if CanClose and IsWindow(helpWindow) then
  SendMessage(helpWindow,wm_close,0,0);
    // HH.HtmlHelp(0, nil, HH_CLOSE_ALL, 0);
end;

procedure TForm1.chkRisingAverageClick(Sender: TObject);
begin
populateUpDownLists;
end;

procedure TForm1.chkRecentCrossingClick(Sender: TObject);
begin
if not hintDown(Sender) then
  populateUpDownLists
else
  chkRecentCrossing.Checked:=not chkRecentCrossing.Checked;
end;

procedure TForm1.chkShowGridClick(Sender: TObject);
begin
if not hintDown(Sender) then
  begin
  if sharesListed then refreshGraph;
  end
else
  chkShowGrid.Checked:=not chkShowGrid.Checked;
end;

procedure TForm1.chkHighLowClick(Sender: TObject);
begin
If PageControl2.ActivePage = moversSheet then
updateMoversLists;
end;

procedure TForm1.spnCrossingChange(Sender: TObject);
begin
if processSpinEdit(spnCrossing) then
  begin
  lblCrossing.Caption:=daysLabel(spnCrossing.Text='1');
	if chkRecentCrossing.Checked then populateUpDownLists;
	end;
end;

procedure TForm1.lstBoughtDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
begin
Accept:=(Source is TListBox);
end;

procedure TForm1.btnRemoveClick(Sender: TObject);
begin
if hintDown(Sender) then Exit;
removePortfolioShare(grdPortfolio.Row);
editRow:=grdPortfolio.RowCount;
updatePortfolioTotal;
end;

procedure TForm1.removePortfolioShare(r:Integer);
var
	i,j:Integer;
begin
With grdPortfolio do begin
j:=ColCount;
while r<=RowCount-2 do
    begin
    i:=0;
    Repeat
    	Cells[i,r]:=Cells[i,r+1];
      inc(i);
    Until i=j;
    inc(r);
    end;
RowCount:=RowCount-1;
btnRemove.Enabled:=(RowCount>2);
if (Row=RowCount-1) then
	if RowCount>2 then Row:=RowCount-2 else btnRemove.Enabled:=false;
end;
portfolioAltered;
end;

procedure TForm1.positionDatePicker;
var
  colNo, xPos:Integer;
begin
with grdPortfolio do
  begin
  colNo:=LeftCol;
  xPos:=Left+ColWidths[0]+1;
  While colNo<portfolioCol[6] do
    begin
    inc(xPos,ColWidths[colNo]+1);
    inc(colNo);
    end;
  end;
DateTimePicker1.Left:=xPos;
end;

procedure TForm1.grdPortfolioDragDrop(Sender, Source: TObject; X,
  Y: Integer);
var
  ACol, ARow:Integer;
begin
grdPortfolio.MouseToCell(X,Y,ACol,ARow);
if (ACol=0) and (ARow>0) and (ARow<grdPortfolio.RowCount-1) then
  begin
  if MessageDlg('Are you sure you want to change the name in this row from "'
+grdPortfolio.Cells[0,ARow]+'" to "'+selectedName+'"?',mtConfirmation, [mbYes, mbNo],0) = mrYes then
    addPortfolioShare(ARow);
  end
else
  addPortfolioShare(0);
end;

procedure warningMessage(const S:String);
begin
MessageDlg(S, mtWarning, [mbOK], 0);
end;

procedure errorMessage(const S:String);
begin
MessageDlg(S, mtError, [mbOK], 0);
end;

function avgValuesString:String;
var
  S:String;
begin
if simpleAvg then S:='S' else S:='E';
Result:=S+'('+IntToStr(lowAverage)+','+IntToStr(highAverage)+')';
end;

procedure TForm1.addPortfolioShare(rowNo:Integer);
var
  i:Integer;
  symbol,S:String;
  year,month,day:Word;
begin
if (companyNo<=numDelisted[currMarket]) and (companyNo>0) then
  begin
  warningMessage('You cannot add this share because it has been delisted.');
  Exit;
  end;
With grdPortfolio do
  begin
  if rowNo=0 then rowNo:=RowCount-1;
  if addSymbol then
    Cells[0,rowNo]:=selectedName
  else
    begin
//    company:=companyAndMarketNo(marketNo,selectedName);
//    companyName:=TrimRight(getName(marketNo,company));
    symbol:=TrimRight(getSymbol(currMarket));
    if lstSymbol.ItemIndex=0 then //prefix
      Cells[0,rowNo]:=symbol+symbolSeparator+selectedName
    else
      Cells[0,rowNo]:=selectedName+symbolSeparator+symbol;
    end;
  if rowNo=RowCount-1 then
    begin
    DecodeDate(Date,year,month,day);
    Cells[portfolioCol[6],rowNo]:=IntToStr(day)+'-'
      +Copy(months,3*month-2,3)+'-'+twoDigits(year mod 100);
    Cells[portfolioCol[5],rowNo]:=priceToStr(Abs(yahooPrices[currMarket,companyNo
    ,length(yahooPrices[currMarket,companyNo])-1]));
    Cells[portfolioCol[7],rowNo]:='0';
    Cells[portfolioCol[10],rowNo]:=avgValuesString;
    updatePortfolioColumns(rowNo,currMarket,companyNo);
    RowCount:=RowCount+1;
    inc(rowNo);
    for i:=0 to 10 do Cells[i,rowNo]:='';
    TopRow:=RowCount-VisibleRowCount;
    //following line led to changePortfolioPrice being called
    Row:=RowCount-2;
    Col:=1;
    end
  else
    updatePortfolioColumns(rowNo,currMarket,companyNo);
{  Cells[portfolioCol[5],nextRow]:=priceToStr(price(currMarket
  ,companyAndMarketNo(currMarket,selectedName),length(yahooPrices[currMarket,])-1));}
  end;
updatePortfolioTotal;
btnRemove.Enabled:=True;
portfolioAltered;
end;

procedure TForm1.updatePortfolioGrid;
begin
if not getPortfolioFile(true) then getPortfolioFile(false);
end;

procedure TForm1.updatePortfolioColumns(gridRow,market,nameNo:Integer);
var
  changePercent, numShares, noDays, i, avg1, avg2: Integer;
	purchPrice,changeAbs: Single;
  sign1,sign2:String[1];
  simpleAvgType:Boolean;
begin
With grdPortfolio do
   begin
   //update original value
   purchPrice:=cellValue(Cells[portfolioCol[5],gridRow]);
   numShares:=StrToInt(Cells[portfolioCol[7],gridRow]);
   Cells[portfolioCol[8],gridRow]:=IntToStr(Round(purchPrice*numShares/100));
   if nameNo>=0 then
      begin
      noDays:=length(yahooPrices[market,nameNo]);
      //update current value
      Cells[portfolioCol[9],gridRow]:=IntToStr(Round(Abs(yahooPrices[market,nameNo,noDays-1])*numShares/100));
      //update current price
      Cells[portfolioCol[4],gridRow]:=priceToStr(Abs(yahooPrices[market,nameNo,noDays-1]));
      //update total movement %
      if purchPrice<>0 then
        changePercent:=Round(100*Abs(yahooPrices[market,nameNo,noDays-1])/purchPrice)-100
      else
        changePercent:=0;
      if changePercent>0 then
   	    Cells[portfolioCol[2],gridRow]:='+' + IntToStr(changePercent)
      else
   	    Cells[portfolioCol[2],gridRow]:=IntToStr(changePercent);
      //update day movement %
      if (noDays>=2) and (Abs(yahooPrices[market,nameNo,noDays-2])>0) then
	      begin
        changePercent:=Round(100*Abs(yahooPrices[market,nameNo,noDays-1])/Abs(yahooPrices[market,nameNo,noDays-2]))-100;
        changeAbs:=Abs(yahooPrices[market,nameNo,noDays-1])-Abs(yahooPrices[market,nameNo,noDays-2]);
        if changePercent>0 then sign1:='+' else sign1:='';
        if changeAbs>=0 then sign2:='+' else sign2:='';
        Cells[portfolioCol[1],gridRow]:=sign1 + IntToStr(changePercent) + '% ('
        + sign2 + priceToStr(changeAbs) + ')';
        end
      else
        Cells[portfolioCol[1],gridRow]:=notAvailStr;
      //update whether above or below moving average
      if (Cells[portfolioCol[10],gridRow]=defaultStr) then
        begin
        simpleAvgType:=simpleAvg;
        avg1:=lowAverage;
        avg2:=highAverage;
        end
      else
        simpleAvgType:=extractAvgSettings(Cells[portfolioCol[10],gridRow],avg1,avg2);
      if noDays>=avg2 then
        begin
        if belowAverage(noDays,market,nameNo,simpleAvgType,avg1,avg2) then
          Cells[portfolioCol[3],gridRow]:=' -'
        else
   	      Cells[portfolioCol[3],gridRow]:=' +';
        end
      else
        Cells[portfolioCol[3],gridRow]:=notAvailStr;
      end
   else
      begin
      for i:=1 to 4 do Cells[portfolioCol[i],gridRow]:=notAvailStr;
      Cells[portfolioCol[9],gridRow]:=notAvailStr;
      end;
  end;
end;

function belowAverage(noDays,market,nameNo:Integer; simpleAvgType:Boolean; lowAvg,highAvg:Word):Boolean;
var
  day, limit: Integer;
	average1, average2: Single;
begin
if simpleAvgType then
  begin
  limit:=noDays-lowAvg;
  average1:=0;
  day:=noDays;
  Repeat
    dec(day);
    average1:=average1+Abs(yahooPrices[market,nameNo,day]);
  Until day=limit;
  limit:=noDays-highAvg;
  average2:=average1;
  Repeat
    dec(day);
    average2:=average2+Abs(yahooPrices[market,nameNo,day]);
  Until day=limit;
  Result:=(average1/lowAvg<=average2/highAvg);
  end
else
  begin
  calcExpAverage(market,nameNo,average1,average2);
  Result:=(average1<=average2);
  end;
end;

procedure TForm1.updatePortfolioTotal;
var
  totalPurchVal, totalCurrVal, totalChange, gridRow: Integer;
begin
totalCurrVal:=0;
totalPurchVal:=0;
with grdPortfolio do
  begin
  gridRow:=RowCount-2;
  While gridRow>0 do
	  begin
    if Cells[portfolioCol[9],gridRow]<>notAvailStr then
      begin
      inc(totalPurchVal,StrToInt(Cells[portfolioCol[8],gridRow]));
      inc(totalCurrVal,StrToInt(Cells[portfolioCol[9],gridRow]));
      end;
    dec(gridRow);
    end;
  end;
if totalPurchVal>0 then
  totalChange:=100 * totalCurrVal div totalPurchVal - 100
else
  totalChange:=0;
lblTotal.Caption:='Total Paid '+IntToStr(totalPurchVal)+', Now '
+IntToStr(totalCurrVal)+', Profit ' + IntToStr(totalChange)+'%';
end;

procedure TForm1.portfolioSheetShow(Sender: TObject);
begin
//updatePortfolioGrid;
if grdPortfolio.RowCount>2 then updateRightControls(grdPortfolio.Cells[0,grdPortfolio.Row]);
end;

procedure TForm1.btnAddClick(Sender: TObject);
begin
if not hintDown(Sender) then addPortfolioShare(0);
end;

procedure TForm1.grdPortfolioKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if not btnRemove.Enabled then Exit;//selected row out of range
If (Key=13) then
  begin
  changePortfolioPrice;
  portfolioAltered;
  end
else if (key=46) then
  btnRemoveClick(Sender);
//else if (ssCtrl in Shift) and (Key=67) {Ctrl C} then
  //copy name of currently selected share to clipboard
//  Clipboard.AsText:=TrimRight(grdPortfolio.Cells[0,grdPortfolio.Row]);
end;

procedure TForm1.changePortfolioPrice;
var
   priceString,name,symbol: String;
   shareNo,marketNo:Integer;
begin
if editRow=grdPortfolio.RowCount-1 then Exit;
priceString:=grdPortfolio.Cells[portfolioCol[5],editRow];
if (priceString>'') and (priceString<>'0') then
	begin
  if (grdPortfolio.Cells[portfolioCol[1],editRow]<>'n/a') and (grdPortfolio.Cells[portfolioCol[1],editRow]<>'') then
    begin
//    portfolioAltered;
    extractNameAndSymbol(grdPortfolio.Cells[portfolioCol[0],editRow],name,symbol);
    marketNo:=mktArrayIndex((Pos(symbol,marketSymbols)-1) div nameLength);
    shareNo:=(Pos(name,namesArray[marketNo])-1) div nameLength;
    updatePortfolioColumns(editRow,marketNo,shareNo);
    end;
  end
else
	grdPortfolio.Cells[portfolioCol[5],editRow]:=oldEditValue;
editRow:=grdPortfolio.RowCount-1;
updatePortfolioTotal;
end;

procedure TForm1.grdPortfolioSetEditText(Sender: TObject; ACol,
  ARow: Integer; const Value: String);
var
  i:Integer;
begin
if (ARow<grdPortfolio.RowCount-1) then
  begin
  for i:=5 to 8 do
    if portfolioCol[i]=ACol then
      begin
      editRow:=ARow;
//      portfolioAltered;
      Exit;
      end;
  end;
grdPortfolio.Cells[ACol,ARow]:=oldEditValue;
end;

procedure TForm1.grdPortfolioGetEditText(Sender: TObject; ACol,
  ARow: Integer; var Value: String);
begin
oldEditValue:=Value;
end;

procedure TForm1.spnAverage2Change(Sender: TObject);
begin
if processSpinEdit(spnAverage2) and (sharesListed) then
	begin
  changeSpinValues(true);
  updatePortfolioGrid;
  populateUpDownLists;
  refreshGraph;
  end;
end;

procedure TForm1.btnHelpMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
if (Button=mbLeft) then
  begin
  if (orderId=0) and (ssCtrl in Shift) then
    begin
    frmRegistration:=TfrmRegistration.Create(self);
    frmRegistration.registerSoftware(0);
    end
  else if (ssAlt in Shift) then
    // edit proxy server settings
    frmProxy:=TfrmProxy.Create(self);
  end;
end;

procedure TForm1.spnRecentDaysChange(Sender: TObject);
begin
if processSpinEdit(spnRecentDays) then
  begin
  lblRecentDays2.Caption:=daysLabel(spnRecentDays.Text='1');
	updateRecentIssues;
  end;
end;

procedure TForm1.lstRecentIssuesClick(Sender: TObject);
begin
if not hintDown(Sender) then
  begin
  btnCombine.Enabled:=btnDelete.Enabled;
  updateRightControls(lstRecentIssues.Items[lstRecentIssues.ItemIndex]);
  end;
end;

procedure TForm1.grdPriceTopLeftChanged(Sender: TObject);
begin
if sharesListed then populatePriceGrid;
end;

procedure TForm1.grdPriceSetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: String);
begin
editPrice:=Value;
SetLength(editPrice,DeleteC(editPrice,','));
end;

procedure TForm1.grdPriceKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if (Key=13) then
	begin
  yahooPrices[currMarket,companyNo,grdPrice.Col]:=cellValue(editPrice);
	btnSave.Enabled:=True;
  pricesChanged[currMarket]:=true;
  //populatePriceGrid;
  refreshGraphSelection;
  refreshGraph;
  populateUpDownLists;
  updateMoversLists;
  updateHighLowLists;
  end;
end;

function priceToWord(p:Single):Word;
begin
Result:=Trunc(p)+16384*Round(4*Frac(p));
end;

procedure TForm1.spnHighLowChange(Sender: TObject);
var
  newSpin,error: Integer;
begin
Val(spnHighLow.Text, newSpin, error);
if (error=0) and (newSpin>=0) then updateHighLowLists;
end;

procedure TForm1.updateHighLowLists;
var
  company, day, startDay, market, noDays: Integer;
  high, low, dayPrice,ratio,highDiff,lowDiff: Single;
begin
ratio:=spnHighLow.Value/100;
lstNearLow.Clear;
lstNearHigh.Clear;
for market:=0 to numMarkets-1 do if marketTicked(market) then
  begin
  setSymbol(market);
  for company:=numDelisted[market]+1 to numCompanies(market)-1 do
    begin
    noDays:=length(yahooPrices[market,company]);
    if lstHighLowRange.ItemIndex=6 then
      startDay:=0
    else
      begin
      startDay:=noDays-rangeOptions[lstHighLowRange.ItemIndex];
      if startDay<0 then startDay:=0;
      end;
    if withinScope(market,company) and (yahooPrices[market,company,noDays-1]<>0) then
      begin
      day:=startDay;
      While yahooPrices[market,company,day]=0 do inc(day);
      high:=0.25;
      low:=999999;
      While day<noDays do
        begin
     	  dayPrice:=Abs(yahooPrices[market,company,day]);
        if dayPrice>0 then
          begin
          high:=Max(high,dayPrice);
          low:=Min(low,dayPrice);
          end;
        inc(day);
        end;
      dayPrice:=Abs(yahooPrices[market,company,noDays-1]);
      highDiff:=1 - (dayPrice/high);
      lowDiff:=(dayPrice/low) - 1;
   	  if (lowDiff<=ratio) and (lowDiff<=highDiff) {use closest} then
        lstNearLow.Items.Add(companyName(market,company))
      else if (highDiff<=ratio) then
        lstNearHigh.Items.Add(companyName(market,company));
      end;
    end;
  end;
lblHighs.Caption:='Highs ('+IntToStr(lstNearHigh.Items.Count)+')';
lblLows.Caption:='Lows ('+IntToStr(lstNearLow.Items.Count)+')';
autoSelectHighLowItem;
end;

procedure TForm1.lstNearHighClick(Sender: TObject);
begin
if hintDown(Sender) then Exit;
lstNearLow.ItemIndex:=-1;
updateRightControls(lstNearHigh.Items[lstNearHigh.ItemIndex]);
end;

procedure TForm1.lstNearLowClick(Sender: TObject);
begin
lstNearHigh.ItemIndex:=-1;
updateRightControls(lstNearLow.Items[lstNearLow.ItemIndex]);
end;

procedure TForm1.highLowSheetShow(Sender: TObject);
begin
autoSelectHighLowItem;
end;

procedure TForm1.PageControl2MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
if btnGetResults.Caption='&Cancel' then cancelBacktest;
If btnHint.Down then
	Application.HelpContext(PageControl2.ActivePage.HelpContext);
end;

procedure TForm1.btnUpdateClick(Sender: TObject);
var
  i:Integer;
  year,month,day,wordDate:Word;
//  pricesToDownload: Boolean;
begin
if hintDown(Sender) then Exit;
if orderId=0 then
  TfrmNag.Create(self).ShowModal
else
  begin
  today:=Date;
  Case DayOfWeek(today) of
    1: today:=today-2;
    7: today:=today-1;
    end;
  DecodeDate(today,year,month,day);
  wordDate:=dateToWord(day,month,year);
  SetLength(marketNotUTD,numMarkets);
//  pricesToDownload:=false;
  for i:=0 to numMarkets-1 do
    begin
    marketNotUTD[i]:=marketTicked(i) and ((length(datesArray[i])=0)
     or (datesArray[i,length(datesArray[i])-1]<wordDate));
//    pricesToDownload:=pricesToDownload or marketNotUTD[i];
    end;
//  if pricesToDownload or not FileExists(marketsFileName) then
//    begin
    frmRegistration:=TfrmRegistration.Create(self);
    frmRegistration.downloadPrices;
    if expired then TfrmNag.Create(self).ShowModal;
//    end
//  else
//    MessageDlg('Your prices are already up-to-date.', mtInformation, [mbOk], 0);
  end;
end;

procedure TForm1.GraphPaint(Sender: TObject);
var
  graphWidth,graphHeight:Integer;
begin
with Graph.Canvas do
  begin
  Brush.Color:=clWhite;
  graphWidth:=Graph.Width;
  graphHeight:=Graph.Height;
  FillRect(Rect(0,0,graphWidth,graphHeight));
  MoveTo(0,0);
  LineTo(0,graphHeight);
  MoveTo(graphWidth-1,0);
  LineTo(graphWidth-1,graphHeight);
  end;
plotGraph(true);
end;

procedure TForm1.GraphMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
If not hintDown2(160) and (Button=mbLeft) then
  begin
  mousePressed:=true;
  if X<graphLeftMargin then
    begin
    zeroY:=not zeroY;
    refreshGraph;
    end;
  showGraphTip(X,Y,Shift);
  end;
end;

procedure TForm1.showGraphTip(X,Y:Integer; Shift:TShiftState);
var
   day:Integer;
   boxWidth,boxHeight:Integer;
   S1,S2,S3:String;
begin
if (X>=graphLeftMargin) and (X<Graph.Width-graphRightMargin) and (Y>0) and (Y<Graph.Height-2) then
	begin
  day:=selectionStart+Round((selectionEnd-selectionStart+1)*(X-graphLeftMargin)
  /(Graph.Width-graphLeftMargin-graphRightMargin));
  if day>selectionEnd then exit;
  S2:=priceToStr(graphData[day,0]);
  if (ssCtrl in Shift) and (highAverage>1) then
    begin
    S3:='('+priceToStr(graphData[day,1]);
    if lowAverage>1 then S3:=S3+','+priceToStr(graphData[day,2]);
    S3:=S3+')';
    boxHeight:=42;
    end
  else
    begin
    S3:='';
    boxHeight:=28;
    end;
	day:=length(datesArray[currMarket])-(length(yahooPrices[currMarket,companyNo])-day);
  S1:=IntToStr(datesArray[currMarket,day] and 31)+'-'
  +Copy(months,3*((datesArray[currMarket,day] shr 5) and 15)-2,3)+'-'
  +twoDigits((datesArray[currMarket,day] shr 9) mod 100);
  with Graph.Canvas do
    begin
    boxWidth:=Max(Max(TextWidth(S1),TextWidth(S2)),TextWidth(S3))+4;
    boxTop:=Y+24;
	  boxLeft:=X-16;
    if boxTop+boxHeight>Graph.Height then
      begin
      dec(boxTop,boxHeight+22);
      dec(boxLeft,10);
      end;
    if boxLeft+boxWidth>Graph.Width then boxLeft:=Graph.Width-boxWidth;
    boxBitmap:=TBitmap.Create;
    boxBitmap.Height:=boxHeight;
    boxBitmap.Width:=boxWidth;
    boxBitmap.Canvas.CopyRect(Bounds(0,0,boxWidth,boxHeight),Graph.Canvas,Bounds(boxLeft,boxTop,boxWidth,boxHeight));
    TextOut(boxLeft+(boxWidth-TextWidth(S1)) div 2,boxTop,S1);
    TextOut(boxLeft+(boxWidth-TextWidth(S2)) div 2,boxTop+13,S2);
    TextOut(boxLeft+(boxWidth-TextWidth(S3)) div 2,boxTop+26,S3);
    end;
  end;
end;

function priceToStr(p:Single):String;
var
	f:Integer;
begin
if Frac(p)=0 then
  Result:=IntToStr(Trunc(p))
else if Frac(4*p)<>0 then
  Result:=FloatToStrF(p,ffFixed,7,2)
else
  begin
  f:=Round(4*Frac(Abs(p)));
  if (f<4) then
    begin
	  if Abs(p)>=1 then
      Result:=IntToStr(Trunc(p))+fractions[f]
    else
      Result:=fractions[f];
    end
  else
	  Result:=IntToStr(Trunc(p+1));
  end;
end;

function singleToStr(p:Single):String;
begin
Result:=FloatToStrF(p,ffFixed,7,2);
end;

procedure TForm1.GraphMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
if boxBitmap<>nil then
  begin
  Graph.Canvas.CopyRect(Bounds(boxLeft,boxTop,boxBitmap.Width,boxBitmap.Height)
,boxBitmap.Canvas,Bounds(0,0,boxBitmap.Width,boxBitmap.Height));
  FreeAndNil(boxBitmap);
  end;
mousePressed:=false;
end;

procedure TForm1.GraphMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
if boxBitmap<>nil then
  begin
  Graph.Canvas.CopyRect(Bounds(boxLeft,boxTop,boxBitmap.Width,boxBitmap.Height)
  ,boxBitmap.Canvas,Bounds(0,0,boxBitmap.Width,boxBitmap.Height));
  FreeAndNil(boxBitmap);
  end;
if mousePressed then showGraphTip(X,Y,Shift);
end;

procedure TForm1.chkRisingClick(Sender: TObject);
begin
if not hintDown(Sender) then
  populateUpDownLists
else
  TCheckBox(Sender).Checked:=not TCheckBox(Sender).Checked;
end;

procedure TForm1.btnHintClick(Sender: TObject);
begin
if btnHint.Down then Screen.Cursor:=crHelp else Screen.Cursor:=crDefault;
end;

procedure TForm1.AddtoPortfolioClick(Sender: TObject);
var
  currentList:TListBox;
begin
currentList:=TListBox(PopupMenu1.PopupComponent);
if currentList=nil then currentList:=TListBox(ActiveControl);
if currentList.ItemIndex>=0 then addPortfolioShare(0);
end;

procedure TForm1.grdPortfolioColumnMoved(Sender: TObject; FromIndex,
  ToIndex: Integer);
var
  i, value:Integer;
begin
DateTimePicker1.Hide;
//determine what value was stored in FromIndex column
//this value is now stored in ToIndex column;
value:=1;
while portfolioCol[value]<>FromIndex do inc(value);
portfolioCol[value]:=ToIndex;
if FromIndex>ToIndex then
  begin
  for i:=1 to 10 do
    if (i<>value) and (portfolioCol[i]>=ToIndex) and (portfolioCol[i]<FromIndex) then
      inc(portfolioCol[i]);
  end
else
  begin
  for i:=1 to 10 do
    if (i<>value) and (portfolioCol[i]>FromIndex) and (portfolioCol[i]<=ToIndex) then
      dec(portfolioCol[i]);
  end;
portfolioAltered;
end;

procedure TForm1.grdPortfolioDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  S: String;
  drawrect :trect;
begin
if ARow=0 then
  begin
  S:=(Sender As TStringgrid).Cells[ ACol, ARow ];
  drawrect := rect;
  DrawText((Sender As TStringgrid).canvas.handle,
            Pchar(S), Length(S), drawrect,
            dt_calcrect or dt_wordbreak or dt_center );
  drawrect.Right := rect.right;
  (Sender As TStringgrid).canvas.fillrect( drawrect );
  DrawText((Sender As TStringgrid).canvas.handle,
            Pchar(S), Length(S), drawrect,
            dt_wordbreak or dt_center);
  end;
end;

procedure TForm1.MonthCalendar1Click(Sender: TObject);
begin
grdPortfolio.Cells[portfolioCol[6],editRow]:=DateToStr(DateTimePicker1.Date);
DateTimePicker1.Hide;
end;

procedure TForm1.grdPortfolioTopLeftChanged(Sender: TObject);
begin
DateTimePicker1.Hide;
editRow:=grdPortfolio.RowCount-1;
end;

procedure TForm1.DateTimePicker1Change(Sender: TObject);
var
  year,month,day:Word;
begin
DecodeDate(DateTimePicker1.Date,year,month,day);
grdPortfolio.Cells[portfolioCol[6],editRow]:=IntToStr(day and 31)+'-'
+Copy(months,3*month-2,3)+'-'+twoDigits(year mod 100);
portfolioAltered;
end;

procedure TForm1.grdPortfolioRowMoved(Sender: TObject; FromIndex,
  ToIndex: Integer);
var
  i,j:Integer;
begin
DateTimePicker1.Hide;
with grdPortfolio do
  if ToIndex=RowCount-1 then
    begin
    for i:=0 to 10 do
      begin
      Cells[i,ToIndex-1]:=Cells[i,ToIndex];
      Cells[i,ToIndex]:='';
      end
    end
  else if FromIndex=RowCount-1 then
    begin
    for i:=ToIndex to FromIndex do
      for j:=0 to 10 do
        Cells[j,i]:=Cells[j,i+1];
    for j:=0 to 10 do
      Cells[j,FromIndex]:='';
    end
  else
    begin
    if (editRow=FromIndex) then
      editRow:=ToIndex
    else if ((FromIndex>ToIndex) and (editRow>=ToIndex) and (editRow<FromIndex)) then
      inc(editRow)
    else if ((FromIndex<ToIndex) and (editRow>FromIndex) and (editRow<=ToIndex)) then
      dec(editRow);
    end;
portfolioAltered;
end;

procedure portfolioAltered;
begin
Form1.btnSave.Enabled:=true;
//Form1.btnPortfolioUndo.Enabled:=true;
portfolioChanged:=true;
end;

procedure TForm1.grdPortfolioMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: Integer;
  dummy :Boolean;
begin
if hintDown(Sender) then Exit;
with grdPortfolio do begin
MouseToCell(X, Y, ACol, ARow);
if (Button=mbRight) and (ARow>0) and (ARow<RowCount-1) then Row:=ARow;
if (ARow=0) then
  begin
  btnRemove.Enabled:=false;
  if Abs(X-ColWidths[0])<5 then portfolioMouseDown:=true;
  if not DateTimePicker1.Visible then changePortfolioPrice;
  end
else
  grdPortfolioSelectCell(Sender,ACol,ARow,dummy);
end;
end;

procedure TForm1.copyPortfolioRow(fromRow,toRow:Integer);
var
  i:Integer;
begin
with grdPortfolio do
  for i:=0 to 9 do Cells[i,toRow]:=Cells[i,fromRow];
end;

procedure TForm1.grdPortfolioDragOver(Sender, Source: TObject; X,
  Y: Integer; State: TDragState; var Accept: Boolean);
begin
Accept:=(Source is TListBox) and (btnAdd.Enabled);
end;

procedure TForm1.grdPortfolioExit(Sender: TObject);
begin
if not(DateTimePicker1.Visible) then changePortfolioPrice;
AddToPortfolio.Visible:=true;
end;

function twoDigits(dateComponent:Word):String;
begin
twoDigits:=StringOfChar('0',Ord(dateComponent<10))+IntToStr(dateComponent);
end;

procedure TForm1.optSimpleClick(Sender: TObject);
begin
hintDown(Sender);
if sharesListed then
  begin
  simpleAvg:=true;
  populateUpDownLists;
  refreshGraph;
  end;
end;

procedure TForm1.ChangeColourClick(Sender: TObject);
var
  clickedLabel: TLabel;
begin
with ColorDialog1 do
  begin
  clickedLabel:=TLabel(ColourMenu.PopupComponent);
  Color:=clickedLabel.Font.Color;
  if Execute and (Color<>clickedLabel.Color){font differs from background} then
    clickedLabel.Font.Color:=Color;
  end;
plotGraph(true);
end;

procedure TForm1.ReturntodefaultClick(Sender: TObject);
var
  clickedLabel: TLabel;
begin
clickedLabel:=TLabel(ColourMenu.PopupComponent);
if clickedLabel.Name='lblAvg1' then
  clickedLabel.Font.Color:=clRed
else
  clickedLabel.Font.Color:=clBlue;
plotGraph(true);
end;

procedure TForm1.Copytoclipboard1Click(Sender: TObject);
var
  currentList:TListBox;
begin
if PopupMenu1.PopupComponent=grdPortfolio then
  Clipboard.AsText:=TrimRight(grdPortfolio.Cells[0,grdPortfolio.Row])
else
  begin
  currentList:=TListBox(PopupMenu1.PopupComponent);
  if currentList=nil then currentList:=TListBox(ActiveControl);
  if currentList.ItemIndex>=0 then
    Clipboard.AsText:=TrimRight(Copy(currentList.Items[currentList.ItemIndex],1,nameLength));
  end;
end;

procedure TForm1.btnPortfolioUndoClick(Sender: TObject);
begin
if not hintDown(Sender) and not getPortfolioFile(portfolioChanged) then
  errorMessage('The required portfolio file is damaged or cannot be found.');
end;

procedure TForm1.btnClearClick(Sender: TObject);
var
  mktName:String;
  i:Integer;
begin
i:=lstMarkets.ItemIndex;
mktName:=TrimRight(Copy(marketSymbols,i*nameLength + 1,nameLength));
if (MessageDlgPos('Are you sure you want to clear all data from the '
+mktName+' market?',mtConfirmation, [mbYes, mbNo], 0, 10, 150) = mrYes) then
  begin
  DeleteFile(mktName+'.mkt');
  DeleteFile(mktName+'.bak');
  if lstMarkets.Checked[i] then
    begin
    lstMarkets.Checked[i]:=false;
    lstMarkets.ItemIndex:=-1;
    pricesChanged[mktArrayIndex(i)]:=false;
    lstMarketsClickCheck(Sender);
    end;
  end;
end;

procedure TForm1.cancelBacktest;
begin
btnGetResults.Caption:='&Start';
lowAverage:=1;
highAverage:=2;
Memo1.Lines.Add('Backtest cancelled.'#13#10);
end;

procedure TForm1.btnGetResultsClick(Sender: TObject);
begin
if hintDown(Sender) then Exit;
if orderId=0 then
  TfrmNag.Create(self).ShowModal
else if sharesListed then
  begin
  if btnGetResults.Caption='&Cancel' then
    cancelBacktest
  else
    begin
    if chkClearResults.Checked then Memo1.Clear;
    if chkOptimize.Checked then
      findProfits
    else
      begin
      Image1.Hide;
      Resize;
      getResults(true);
      end;
    end;
  end;
end;

function TForm1.getResults(updateMemo:Boolean):Single;
var
  day, lastTradeDay, arrayIndex,error,dateOffset: Integer;
  longestHeld,shortestHeld,winningPositions,IMR,maxBet,spread, noPositions:Integer;
  totalProfit:Single;
  biggestWin, biggestLoss, profit,marginCall,bet,startBet: Single;
  above,spreadBet,goShort,goLong,positionClosed,listTrades:Boolean;
  S:String;
begin
listTrades:=updateMemo and chkListTrades.Checked;
spreadBet:=chkSpreadBet.Checked;
goShort:=lstTradeType.ItemIndex>0;
goLong:=lstTradeType.ItemIndex<>1;
noPositions:=0;
winningPositions:=0;
if spreadBet then
  begin
  totalProfit:=0;
  biggestWin:=0;
  biggestLoss:=0;
  maxBet:=spnMaxBet.Value;
  spread:=spnSpread.Value;
  Val(txtMinBet.Text, startBet, error);
  if (error>0) then
    begin
    warningMessage('Invalid minimum bet specified. Defaulting to 1.');
    startBet:=1;
    txtMinBet.Text:='1';
    end
  else if startBet>maxBet then
    begin
    warningMessage('Minimum bet more than maximum. Defaulting to same as maximum.');
    startBet:=maxBet;
    txtMinBet.Text:=IntToStr(maxBet);
    end;
  bet:=startBet;
  IMR:=spnMargin.Value;
  marginCall:=0;
  end
else
  begin
  totalProfit:=1;
  biggestWin:=1;
  biggestLoss:=1;
  end;
longestHeld:=0;
lastTradeDay:=0;
positionClosed:=false;
if updateMemo then
  begin
  if chkSpreadbet.Checked then
    S:='Spread betting'
  else
    S:='Trading';
  dateOffset:=length(datesArray[currMarket])-length(yahooPrices[currMarket,companyNo]);
  S:=S+' "'+TrimRight(selectedName)+'" from '+ dayToStr(dateOffset+selectionStart)
  +' to '+ dayToStr(dateOffset+selectionEnd)+', ';
  if goShort and goLong then
    S:=S+'going both long and short'
  else if goLong then
    S:=S+'going long only'
  else
    S:=S+'going short only';
  S:=S+', using ';
  if lowAverage>1 then
    S:=S+IntToStr(lowAverage)+'-day and '+IntToStr(highAverage)
  else
    S:=S+'a single '+IntToStr(highAverage);
  if optTestSimple.Checked then
    S:=S+'-day simple moving average'
  else
    S:=S+'-day exponential moving average';
  if lowAverage>1 then S:=S+'s';
  S:=S+'.';
  if listTrades then S:=S+#13#10#13#10'Trades executed (period held & profit):'#13#10;
  Memo1.Lines.Add(S);
  end;
day:=getFirstDay(highAverage);
if day=0 then day:=1;//following calculations use day-1
shortestHeld:=selectionEnd;
if lowAverage=1 then arrayIndex:=0 else arrayIndex:=1;
//if day<selectionEnd then
if day<selectionEnd then
  begin
  above:=(graphData[day,arrayIndex]>graphData[day,arrayIndex+1]);
  While day<selectionEnd do
    begin
    inc(day);
    if above and (graphData[day,arrayIndex]<graphData[day,arrayIndex+1]) then
      begin
      above:=false; //low average fallen below high average
      // if long position is open, close it
      if (lastTradeDay>0) and (graphData[lastTradeDay,0]>0) and goLong then
        positionClosed:=true
      else
        lastTradeDay:=day;//start of first trade
      end
    else if (not above) and (graphData[day,arrayIndex]>graphData[day,arrayIndex+1]) then
      begin
      above:=true;
      //if short position is open, close it
      if (lastTradeDay>0) and (graphData[lastTradeDay,0]>0) and goShort then
        positionClosed:=true
      else
        lastTradeDay:=day;
      end;
    if positionClosed then //open position was closed
      begin
      positionClosed:=false;
      inc(noPositions);
      if (day-lastTradeDay)>longestHeld then
        longestHeld:=day-lastTradeDay
      else if (day-lastTradeDay)<shortestHeld then
        shortestHeld:=day-lastTradeDay;
      if spreadBet then
        begin
        //use points difference
        if above then //closed short position
          profit:=bet*(graphData[lastTradeDay,0]-graphData[day,0]-spread)
        else //closed long position
          profit:=bet*(graphData[day,0]-graphData[lastTradeDay,0]-spread);
        totalProfit:=totalProfit+profit;
        end
      else
        begin
        //use price ratio
        if above then
          profit:=(graphData[lastTradeDay,0]/graphData[day,0])
        else
          profit:=(graphData[day,0]/graphData[lastTradeDay,0]);
        totalProfit:=totalProfit*profit;
        end;
      if listTrades then
        begin
        S:=dayToStr(dateOffset+lastTradeDay)+' - ' + dayToStr(dateOffset+day)+': ';
        if spreadBet then
          Form1.Memo1.Lines.Add(S+singleToStr(profit) + ' ('+ singleToStr(bet)+'/pt)')
        else
          Form1.Memo1.Lines.Add(S+percentStr(profit-1));
        end;
      if spreadBet then //separate branch as previous bet value used for memo update
        begin
        if totalProfit>=0 then
          begin
          bet:=startBet+totalProfit/IMR;
          if bet>maxBet then bet:=maxBet;
          end
        else
          begin
          marginCall:=marginCall-totalProfit;
          totalProfit:=0;
          bet:=startBet;
          end;
        end;
      if (spreadBet and (profit>0)) or (profit>1) then
        begin
        inc(winningPositions);
        if profit>biggestWin then biggestWin:=profit;
        end
      else if (profit<biggestLoss) then
        biggestLoss:=profit;
      if (above and goLong) or (not above and goShort) then lastTradeDay:=day;
      end;
    end;
  end;
if noPositions=0 then
  Result:=0
else if spreadBet then
  Result:=totalProfit-marginCall
else
  Result:=totalProfit-1;
if updateMemo then with Memo1.Lines do
  begin
  if noPositions>0 then
    begin
    Add(#13#10'No. of positions: '+IntToStr(noPositions)+' ('+IntToStr(Round(100*winningPositions/noPositions))+'% winning)');
    if spreadBet then
      begin
      Add('Total winnings: '+ singleToStr(totalProfit));
      Add('Margin call total: '+singleToStr(marginCall));
      Add('Biggest win: '+ singleToStr(biggestWin));
      Add('Biggest loss: '+ singleToStr(Abs(biggestLoss)));
      end
    else
      begin
      Add('Biggest gain: '+ percentStr(biggestWin-1));
      Add('Biggest loss: '+ percentStr(1-biggestLoss));
      Add('Total profit: '+percentStr(Result));
      end;
    if longestHeld>0 then Add('Longest time held: '+IntToStr(longestHeld)+' '+daysLabel(longestHeld=1));
    if shortestHeld<selectionEnd then Add('Shortest time held: '+IntToStr(shortestHeld)+' '+daysLabel(shortestHeld=1));
    end
  else
    Add('No trades possible.');
  Add('----------------------------------------------------');
  end;
end;

function daysLabel(singular:Boolean):String;
begin
if singular then Result:='day' else Result:='days';
end;

function percentStr(i:Single):String;
begin
Result:=FloatToStrF(100*i,ffFixed,7,2)+'%';
end;

procedure TForm1.spnTestAv1Change(Sender: TObject);
begin
if (btnGetResults.Caption='&Start') and (sharesListed) and processSpinEdit(spnTestAv1) then
  begin
  changeSpinValues(false);
  refreshGraph;
  end;
end;

procedure TForm1.spnTestAv2Change(Sender: TObject);
begin
if (btnGetResults.Caption='&Start') and (sharesListed) and processSpinEdit(spnTestAv2) then
  begin
  changeSpinValues(false);
  refreshGraph;
  end;
end;

procedure TForm1.optExpClick(Sender: TObject);
begin
hintDown(Sender);
if sharesListed then
  begin
  simpleAvg:=false;
  populateUpDownLists;
  refreshGraph;
  end;
end;

procedure TForm1.optTestExpClick(Sender: TObject);
begin
hintDown(Sender);
if sharesListed then
  begin
  simpleAvg:=false;
  refreshGraph;
  end;
end;

procedure TForm1.findProfits;
type
  TRGBTripleArray = array[Word] of TRGBTriple;
  pRGBTripleArray = ^TRGBTripleArray;
var
  bestLowAverage,bestHighAverage,i:Integer;
  range,maxProfit,minProfit,profit:Single;
  shade:Byte;
  bitmap:TBitmap;
  row:pRGBTripleArray;
begin
btnGetResults.Caption:='&Cancel';
if chkClearResults.Checked then Memo1.Clear;
maxProfit:=-1e38;
minProfit:=1e38;
bestLowAverage:=0;
bestHighAverage:=0;
if selectionStart>selectionEnd-198 then lowAverage:=selectionEnd-selectionStart-1 else lowAverage:=198;
spnTestAv1.Value:=1;
Repeat
spnTestAv2.Value:=lowAverage;
Application.ProcessMessages;
if btnGetResults.Caption='&Start' then exit;
if selectionStart>selectionEnd-199 then highAverage:=selectionEnd-selectionStart else highAverage:=199;
  Repeat
    calculateGraph;
    profit:=getResults(false);
    contourData[lowAverage,highAverage]:=profit;
    contourData[highAverage,lowAverage]:=profit;
    if profit>maxProfit then
      begin
      bestLowAverage:=lowAverage;
      bestHighAverage:=highAverage;
      maxProfit:=profit;
      end
    else if profit<minProfit then
      minProfit:=profit;
    dec(highAverage);
  Until highAverage<=lowAverage;
dec(lowAverage);
Until lowAverage=0;
//if btnGetResults.Caption='&Start' then exit;
btnGetResults.Caption:='&Start';
//normalise contour data
bitmap:=TBitmap.Create;
with bitmap do
  begin
  Width:=241;
  Height:=241;
  PixelFormat:=pf24bit;
  end;
//draw scales
With bitmap.Canvas do
  begin
  Font.Name:='MS Serif';
  Font.Size:=7;
  //y-axis
  MoveTo(25,3);
  LineTo(25,203);
  //x-axis
  MoveTo(25,203);
  LineTo(225,203);
  //top border
  MoveTo(25,3);
  LineTo(225,3);
  //right border
  MoveTo(225,3);
  LineTo(225,203);
  i:=3;
  Repeat
    MoveTo(25,i);
    LineTo(20,i);
    TextOut(5,i-5,IntToStr(203-i));
    inc(i,20);
  Until i=223;
  i:=25;
  Repeat
    MoveTo(i,203);
    LineTo(i,208);
    TextOut(i-5,208,IntToStr(i-25));
    inc(i,20);
  Until i=245;
  end;
range:=maxProfit-minProfit;
if range>0 then
  begin
  if selectionStart>selectionEnd-199 then lowAverage:=selectionEnd-selectionStart else lowAverage:=199;
  Repeat
    if selectionStart>selectionEnd-199 then highAverage:=selectionEnd-selectionStart else highAverage:=199;
    row:=bitmap.ScanLine[203-lowAverage];
    Repeat
      shade:=Round((contourData[lowAverage,highAverage]-minProfit)*255 / range);
      with row[highAverage+25] do
        begin
        rgbtRed:=shade;
        rgbtGreen:=shade;
        rgbtBlue:=shade;
        end;
      dec(highAverage);
    Until highAverage=0;
  dec(lowAverage);
  Until lowAverage=0;
  end;
Image1.Picture.Graphic:=bitmap;
bitmap.Free;
Image1.Show;
lowAverage:=bestLowAverage;
highAverage:=bestHighAverage;
Resize;
Memo1.Lines.Add('Best results from optimization backtest:'#13#10);
spnTestAv1.Value:=bestLowAverage;
spnTestAv2.Value:=bestHighAverage;
getResults(true);
//changeSpinValues(false);
//refreshGraph;
end;

procedure TForm1.Image1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
If (Button=mbLeft) and not hintDown2(620) then
	begin
  if (X<26) or (X>225) or (Y<4) or (Y>202) then Exit;
  lblContourTip.Show;
  Screen.Cursor:=crCross;
  updateGauge(X,Y);
  end;
end;

procedure TForm1.Image1MouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
lblContourTip.Hide;
Screen.Cursor:=crDefault;
end;

procedure TForm1.Image1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
if (X<26) or (X>225) or (Y<4) or (Y>202) then
  lblContourTip.Hide
else if Screen.Cursor=crCross then
  begin
  lblContourTip.Show;
  updateGauge(X,Y);
  end;
end;

procedure TForm1.updateGauge(X,Y:Integer);
var
  S1,S2:String;
begin
with lblContourTip do
  begin
  Left:=Image1.Left+X-Width div 2;
  if Y>30 then
    Top:=Image1.Top+Y-35
  else
    Top:=Image1.Top+Y+10;
  Y:=203-Y;
  dec(X,25);
  S1:=percentStr(contourData[Y,X]);
  S2:='('+IntToStr(X)+','+IntToStr(Y)+')';
  Width:=Max(Canvas.TextWidth(S1),Canvas.TextWidth(S2));
  Height:=2*Canvas.TextHeight(S1+S2);
  Caption:=S1+#13#10+S2;
  end;
if X<>spnTestAv2.Value then spnTestAv2.Value:=X;
if Y<>spnTestAv1.Value then spnTestAv1.Value:=y;
end;

procedure TForm1.optTestSimpleClick(Sender: TObject);
begin
if sharesListed then
  begin
  simpleAvg:=true;
  refreshGraph;
  end;
end;

procedure TForm1.grdPortfolioMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow, blankRow, lastRow, startRow, nextRow, i, cellType: Integer;
  S: String;
begin
positionDatePicker;
with grdPortfolio do begin
MouseToCell(X, Y, ACol, ARow);
if (ARow=0) then
  begin
  portfolioAltered;
  If (Button=mbRight) then
    begin
    blankRow:=RowCount-1;
    cellType:=0;
    while portfolioCol[cellType]<>ACol do inc(cellType);
    lastRow:=blankRow-1;
    //sort rows according to column clicked
    startRow:=1;
    While startRow<lastRow do
      begin
      S:=Cells[ACol,startRow];//comparison string
      nextRow:=startRow;
      i:=startRow;
      While i<lastRow do
        begin
        inc(i);
        if (reverseSort xor isCellLower(cellType,Cells[ACol,i],S)) then
          begin
          S:=Cells[ACol,i];
          nextRow:=i;
          end;
        end;
      //swap startRow with nextRow, using last blank row of grid as temp store
      copyPortfolioRow(startRow,blankRow);//copy startRow to blank row
      copyPortfolioRow(nextRow,startRow);
      copyPortfolioRow(blankRow,nextRow);
      inc(startRow);
      end;
    //make last row blank again
    for i:=0 to 10 do Cells[i,blankRow]:='';
    reverseSort:=not reverseSort;
    end
  else
    portfolioMouseDown:=false;
  end;
end;
end;

function TForm1.isCellLower(colType:Integer; const cellStr,compStr:String):Boolean;
begin
if cellStr=notAvailStr then
  Result:=false
else if compStr=notAvailStr then
  Result:=true
else Case colType of
  0: Result:=(LowerCase(cellStr)<LowerCase(compStr)); //sharename
  1: Result:=(percentChange(cellStr)<percentChange(compStr));
  3: Result:=(cellStr>compStr);//Av Pos
  4,5: Result:=(cellValue(cellStr)<cellValue(compStr));//price now or paid
  6: Result:=(dateStrToWord(cellStr)<dateStrToWord(compStr));//date of purchase
  else Result:=(StrToInt(cellStr)<StrToInt(compStr)); //integers
  end;
end;

function percentChange(const changeStr:String):Integer;
var
  i:Integer;
begin
i:=1;
while changeStr[i]<>'%' do inc(i);
Result:=StrToInt(Copy(changeStr,1,i-1));
end;

procedure TForm1.spnThresholdChange(Sender: TObject);
begin
if processSpinEdit(spnThreshold) then
  begin
  threshold:=spnThreshold.Value;
  if (scopeSetting and 2)>0 then refreshLists;
  end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
var
   serialNum,defaults,xy,i:Integer;
   dummy:Cardinal;
begin
boxBitmap.Free;
if (orderId=0) or expired then TfrmNag.Create(self).ShowModal;
With TRegistry.Create do
  begin
	try
		RootKey:=HKEY_CURRENT_USER;
		if OpenKey(regKey,true) then
			begin
    	defaults:=Ord(Form1.WindowState=wsMaximized) or (spnAverage1.Value shl 1) or (spnAverage2.Value shl 10)
    		or (spnCrossing.Value shl 19) or (lstRange.ItemIndex shl 26);
			WriteInteger('defaults',defaults);
    	xy:=Form1.Width or (Form1.Height shl 16);
			WriteInteger('xy',xy);
			GetVolumeInformation(PChar(Copy(Application.ExeName,1,3)), nil, 0, @serialNum, dummy, dummy, nil, 0);
    	WriteInteger('flags',lstScope.ItemIndex or (Ord(chkScope.Checked) shl 1)
      or (Ord(chkIndexOnly.Checked) shl 2)
      or (Ord(chkRising.Checked) shl 3)
      or (Ord(chkFalling.Checked) shl 4)
      or (Ord(chkRecentCrossing.Checked) shl 5)
			or (Ord(chkShowGrid.Checked) shl 6)
    	or (Ord(optSimple.Checked) shl 7)
    	or (spnHighLow.Value shl 8)
    	or (((orderId xor serialNum xor xy xor defaults) and 131071) shl 14));
    	WriteInteger('user',orderId);
      WriteInteger('colours',ColorToCode(lblAvg1.Font.Color) or (ColorToCode(lblAvg2.Font.Color) shl 6)
      or (lstHighLowRange.ItemIndex shl 12) or (threshold shl 18));
      WriteInteger('backtest',Ord(optTestSimple.Checked) or (Ord(chkOptimize.Checked) shl 1) 
      or (Ord(chkListTrades.Checked) shl 2) or (Ord(chkSpreadBet.Checked) shl 3)
      or (Ord(chkClearResults.Checked) shl 4) or (Ord(chkUnadjusted.Checked) shl 5)
      or (lstTradeType.ItemIndex shl 6)
      or (spnTestAv1.Value shl 8) or (spnTestAv2.Value shl 17));
      WriteInteger('betoptions',spnMargin.Value or (spnMaxBet.Value shl 16)
       or (Ord(chkSymbol.Checked) shl 25) or (lstSymbol.ItemIndex shl 26));
      WriteString('minbet',txtMinBet.Text);
      i:=lstPortfolio.ItemIndex;
      if (i<0) or (grdPortfolio.RowCount=2) then i:=0;
      WriteInteger('betoptions2', spnSpread.Value or (i shl 10));
      if mainUrl<>defaultUrl then WriteString('url',mainUrl);
    	end;
	except end;
  Free;
  end;
end;

procedure TForm1.portfolioSheetMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
If (btnHint.Down) then
	Application.HelpContext(PageControl2.ActivePage.HelpContext);
end;

procedure TForm1.PageControl2Change(Sender: TObject);
begin
if (sharesListed) then
  begin
  changeSpinValues(PageControl2.ActivePageIndex<>5);
  refreshGraph;
  if PageControl2.ActivePageIndex<>3 then normalPortfolioSize;
  end;
end;

procedure TForm1.normalPortfolioSize;
begin
PageControl2.Width:=249;
portfolioSheet.Width:=241;
grdPortfolio.Width:=234;
lblTotal.Width:=234;
btnResize.Left:=218;
btnResize.caption:='>';
btnAdd.Left:=3;
btnRemove.Left:=82;
btnPortfolioUndo.Left:=161;
lblPortfolio.Left:=8;
lstPortfolio.Left:=62;
Resize;
end;

procedure TForm1.btnResizeClick(Sender: TObject);
begin
if hintDown(Sender) then Exit;
DateTimePicker1.Hide;
if btnResize.caption='>' then
  begin
  PageControl2.Width:=465;
  portfolioSheet.Width:=458;
  grdPortfolio.Width:=451;
  lblTotal.Width:=451;
  btnResize.caption:='<';
  btnResize.Left:=435;
  btnAdd.Left:=115;
  btnRemove.Left:=194;
  btnPortfolioUndo.Left:=273;
  lblTotal.Top:=328;
  lblPortfolio.Left:=126;
  lstPortfolio.Left:=180;
  Resize;
  end
else
  normalPortfolioSize;
autoSelectPortfolioItem;
end;

function TForm1.hintDown(Sender:TObject):Boolean;
begin
Result:=btnGetResults.Caption='&Cancel';
if Result then
  cancelBacktest
else
  begin
  Result:=btnHint.Down;
  if Result then
    begin
    helpWindow:=HtmlHelp(GetDesktopWindow, '..\AutoShare.chm', HH_HELP_CONTEXT, TWinControl(Sender).HelpContext);
    btnHint.Down:=false;
    Screen.Cursor:=crDefault;
    end;
  end;
end;

function TForm1.hintDown2(helpId: THelpContext):Boolean;
begin
Result:=btnHint.Down;
if Result then
  begin
  helpWindow:=HtmlHelp(GetDesktopWindow, '..\AutoShare.chm', HH_HELP_CONTEXT, helpId);
  btnHint.Down:=false;
  Screen.Cursor:=crDefault;
  end;
end;

procedure TForm1.lstNearLowMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
hintDown(Sender);
end;

procedure TForm1.printHeadings(pColWidths: array of Integer);
var
  x, col:Integer;
begin
x:=150;
with Printer.Canvas do
  begin
  Font.Style:=[fsUnderline];
  TextOut(x,350,'Share name/symbol');
  for col:=1 to 9 do
    begin
    inc(x,pColWidths[col-1]+50);
    TextOut(x,350,grdPortfolio.Cells[col,0]);
    end;
  Font.Style:=[];
  end;
end;

procedure TForm1.printPortfolioClick(Sender: TObject);
var
  col, row, x, y, vSpacing: Integer;
  pColWidths: array[0..9] of Integer;
begin
if not PrintDialog1.Execute then Exit;
Application.ProcessMessages;//close print dialog immediately
with Printer do
  begin
  Title:='Portfolio "'+lstPortfolio.Text+'" on '+DateToStr(Date);
  BeginDoc;
  with Canvas do
    begin
    Font.Name:='Times New Roman';
    Font.Size:=14;
    Pen.Color:=clBlack;
    Font.Style:=[fsUnderline];
    vSpacing:=TextHeight('W');
    TextOut((PageWidth-TextWidth(Title)) div 2, 100, Title);
    Font.Size:=10;
    for col:=1 to 9 do
      pColWidths[col]:=Canvas.TextWidth(grdPortfolio.Cells[col,0]);
    pColWidths[0]:=Canvas.TextWidth('Share name/symbol');
    end;
  printHeadings(pColWidths);
  y:=500;
  for row:=1 to grdPortfolio.RowCount - 2 do
    begin
    x:=150;
    for col:=0 to 9 do
      begin
      Canvas.TextOut(x + (pColWidths[col]-Canvas.TextWidth(grdPortfolio.Cells[col,row])) div 2
      ,y,grdPortfolio.Cells[col,row]);
      inc(x,pColWidths[col]+50);
      end;
    inc(y,vSpacing);
    if y+100>PageHeight then
      begin
      NewPage;
      printHeadings(pColWidths);
      y:=500;
      end;
    end;
  EndDoc;
  end;
end;

procedure TForm1.exportPortfolioClick(Sender: TObject);
var
  F:TextFile;
  row,col:Integer;
  S:String;
begin
with SaveDialog1 do
  begin
  FileName:=lstPortfolio.Text;
  Title:='Export Portfolio';
  Filter := '*.csv|*.csv';
  if not Execute then Exit;
  end;
try
  AssignFile(F,SaveDialog1.FileName);
  Rewrite(F);
  for row:=0 to grdPortfolio.RowCount-2 do
    begin
    S:='';
    for col:=0 to 9 do
      if (row>0) and ((col=portfolioCol[4]) or (col=portfolioCol[5])) then
        S:=S+','+priceToStr(cellValue(grdPortfolio.Cells[col,row]))
      else
        S:=S+','+grdPortfolio.Cells[col,row];
    WriteLn(F,Copy(S,2,65535));
    end;
  CloseFile(F);
except
  errorMessage('Could not save portfolio. Please check that the target file is not already open.');
  end;
end;

procedure TForm1.lstPortfolioChange(Sender: TObject);
var
  i:Integer;
begin
with lstPortfolio do
  begin
  i:=ItemIndex;
  if (i<0) then Exit;
  if Text='' then
    begin
    Text:=Items[i];//reinstate deleted text
    updatePortfolioGrid;
    Exit;
    end;
  if Text<>Items[i] then exit;
  if portfolioChanged and
  (MessageDlgPos('Save changes to '''+Items[editedIndex]
  +'''?',mtConfirmation, [mbYes, mbNo], 0, Left+261, Top+132) = mrYes) then
    savePortfolio;
  if i=Items.Count-1 then//[Add new portfolio] entry at bottom chosen
    newPortfolio
  else
    updatePortfolioGrid;
  editedIndex:=i;
  end;
end;

procedure TForm1.newPortfolio;
var
  i:Integer;
begin
with grdPortfolio do
  begin
  //new portfolio
  RowCount:=2;
  RowHeights[0]:=2*Canvas.TextHeight('W')+1;
  for i:=0 to 10 do
    begin
    portfolioCol[i]:=i;
    ColWidths[i]:=columnWidths[i];
    Cells[portfolioCol[i],0]:=columnLabels[i];
    Cells[i,1]:='';
    end;
  btnRemove.Enabled:=False;
  updatePortfolioTotal;
  portfolioChanged:=false;
  btnSave.Enabled:=anyPricesChanged;
  btnAdd.Enabled:=false;
  DateTimePicker1.Hide;
  editRow:=1;
  end;
end;

procedure TForm1.lstPortfolioKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
with lstPortfolio do begin
if Key=13 then
  begin
  editPortfolioName;
  DroppedDown:=true;
  end
else if (Key=46) and (SelText=Text) and (ItemIndex<>Items.Count-1) and
(MessageDlgPos('Are you sure you want to delete the portfolio '''
+lstPortfolio.Items[ItemIndex]+'''?',mtConfirmation, [mbYes, mbNo], 0, 10, 150) = mrYes) then
  begin
  DeleteFile(Text+'.dat');
  DeleteFile(Text+'.bak');
  Items.Delete(ItemIndex);
  ItemIndex:=0;
  portfolioChanged:=false;
  end;
end;
end;

procedure TForm1.editPortfolioName;
var
  itemName, oldName:String;
  i:Integer;
begin
with lstPortfolio do
  begin
  itemName:=Text;
  oldName:=Items[editedIndex];
  if (itemName='') or (itemName=oldName) or DroppedDown then Exit;
  //edited item
  if editedIndex=Items.Count-1 then
    begin
    //new portfolio given a name
    i:=Items.IndexOf(itemName);
    if i=-1 then //doesn't already exist
      Items.Add(newPortfolioText)
    else
      begin
      ItemIndex:=i;
      lstPortfolioChange(nil);
      Exit;
      end;
    end
  else
    begin
    RenameFile(oldName+'.dat',itemName+'.dat');
    RenameFile(oldName+'.bak',itemName+'.bak');
    end;
  Items.Strings[editedIndex]:=itemName;
  end;
end;

procedure TForm1.lstRangeChange(Sender: TObject);
begin
if sharesListed then
  begin
  if btnGetResults.Caption='&Cancel' then cancelBacktest;
  if lstRange.ItemIndex<7 then
    begin
    GraphScroll.Visible:=true;
    Graph.Height:=Form1.Height-180;
    end
  else
    begin
    GraphScroll.Visible:=false;
    Graph.Height:=Form1.Height-162;
    end;
  refreshGraphSelection;
  refreshGraph;
  end;
end;

{function priceValue(const priceString: String):Integer;
var
   value, error1, error2: Integer;
begin
Val(priceString, value, error1);
If error1=0 then
   Result:=value
else
   begin
   If error1>1 then
      Val(Copy(priceString,1,error1-1),value,error2);
   //ends in fraction chr
   Case priceString[error1] of
      '¼': Result:=value or 16384;
      '½': Result:=value or 32768;
      '¾': Result:=value or 49152;
   else
   		Result:=value;
   		end;
   end;
end;

procedure savePrice(const nameString,priceString:String; today:Integer);
var
   shareNo,price:Cardinal;
   targetLength: Integer;
   i,n:Integer;
begin
targetLength:=length(namesString);
price:=priceValue(priceString);
if price=0 then Exit;
//find share no
i:=1;
Repeat
n:=1;
While (nameString[n]=namesString[i]) do
  begin
 	inc(i);
  inc(n);
  end;
inc(i,nameLength+1-n);
Until (n>nameLength) or (i>=targetLength);
if n>nameLength then //share exists (only process shares existing today)
	begin
	shareNo:=(i-1) div nameLength - 1;
  pricesArray[shareNo,today]:=price;
  end;
end;

procedure changePrice(const nameString,changeString:String; today:Integer);
var
   shareNo,price:Cardinal;
   targetLength: Integer;
   i,n:Integer;
   fraction,change: Integer;
begin
targetLength:=length(namesString);
//find share no
i:=1;
Repeat
n:=1;
While (nameString[n]=namesString[i]) do
  begin
 	inc(i);
  inc(n);
  end;
inc(i,nameLength+1-n);
Until (n>nameLength) or (i>=targetLength);

if n>nameLength then //share exists (only process shares existing today)
	begin
	shareNo:=(i-1) div nameLength - 1;
  if changeString>'' then
    begin
	  change:=priceValue(changeString);
 	  fraction:=(pricesArray[shareNo,today] and 49152) - (change and 49152);
    price:=(pricesArray[shareNo,today] and 16383) - (change and 16383) + fraction;
    if fraction<0 then inc(price,-1+65536);
    pricesArray[shareNo,today-1]:=price;
    end
  else
    pricesArray[shareNo,today-1]:=pricesArray[shareNo,today];
  end;
end;

procedure convertToFourYears;
var
  fileDayName: String;
  teletextFile: TextFile;
  dateToday: TDate;
  year,month,day:Word;
  teletextLine: String[30];
  i,offset:Integer;
const
  dir: String='C:\Projects\Server\Socket\Converter\Prices\';
begin
//shift prices up by 250 days
{for day:=999 downto 250 do
  begin
  i:=0;
  while i<numCompanies do
   	begin
  	pricesArray[i,day]:=pricesArray[i,day-250];
   	inc(i);
   	end;
  datesArray[day]:=datesArray[day-250];
  end;
for day:=0 to 249 do
  begin
  for i:=0 to numCompanies-1 do pricesArray[i,day]:=0;
  datesArray[day]:=0;
  end;
//collect prices for first year from .txt files
// in C:\Projects\Server\Socket\Converter\Prices
//Exit;
dateToday:=EncodeDate(2002,12,31);
i:=0;
Repeat
  Repeat
	  DecodeDate(dateToday,year,month,day);
	  fileDayName:=twoDigits(year-2000)+twoDigits(month)+twoDigits(day)+'.txt';
    dateToday:=dateToday+1;
  Until FileExists(dir+'teletext'+fileDayName) or FileExists(dir+fileDayName);
  datesArray[i]:=((year-2000) shl 9) or (month shl 5) or (day);
  //process each teletext file
  if FileExists(dir+fileDayName) then
    begin
    offset:=18;
    AssignFile(teletextFile,dir+fileDayName);
    end
  else
    begin
    offset:=14;
    AssignFile(teletextFile,dir+'teletext'+fileDayName);
    end;
  Reset(teletextFile);
  Repeat
    ReadLn(teletextFile,teletextLine);
    If teletextLine[1]<>' ' then
      begin
      savePrice(Copy(teletextLine,1,nameLength),Copy(teletextLine,10,4),i);
      savePrice(Copy(teletextLine,offset,nameLength),Copy(teletextLine,offset+9,4),i);
      end;
  Until Eof(teletextFile);
  CloseFile(teletextFile);
  inc(i);
Until i=250;
Form1.btnSave.Enabled:=True;
pricesChanged:=true;
//correct 031002 problem by using next day's differences
i:=0;
//find array pos for 031002
while datesArray[i]<>((2003-2000) shl 9) or (10 shl 5) or (3) do inc(i);
//load following day's differences
AssignFile(teletextFile,dir+'031003.txt');
Reset(teletextFile);
  Repeat
    ReadLn(teletextFile,teletextLine);
    changePrice(Copy(teletextLine,1,nameLength),Copy(teletextLine,14,4),i);
    changePrice(Copy(teletextLine,18,nameLength),Copy(teletextLine,31,4),i);
  Until Eof(teletextFile);
CloseFile(teletextFile);
end;}


procedure TForm1.lstHighLowRangeChange(Sender: TObject);
begin
updateHighLowLists;
end;

procedure TForm1.chkScopeClick(Sender: TObject);
begin
if hintDown(Sender) then Exit;
scopeSetting:=scopeSetting xor 2;
lstScope.Enabled:=chkScope.Checked;
spnThreshold.Enabled:=chkScope.Checked;
refreshLists;
end;

procedure TForm1.lstScopeChange(Sender: TObject);
begin
scopeSetting:=(scopeSetting and 2) or lstScope.ItemIndex;
refreshLists;
end;

procedure TForm1.processMarketsFile;
var
  P,limit:PChar;
  inBuf,outBuf:Pointer;
  bufferSize,i,start:Integer;
  line,problems,mktName,pricesUrl,indexCountry,currentUrlCountry:String;
  outputSize,closingTime:Word;
  searchResult:TSearchRec;
begin
marketSymbols:='';
currentUrlCountry:='';
if FileExists(marketsFileName) then
  begin
  With TFileStream.Create(marketsFileName,fmOpenRead or fmShareDenyNone) do
    begin
    bufferSize:=Size;
    GetMem(inBuf,bufferSize);
    Read(outputSize,2);
    outputSize:=Swap(outputSize);
    Read(inBuf^,bufferSize);
    Free;
    end;
  try
    DecompressBuf(inBuf,bufferSize,outputSize,outBuf,i);
    FreeMem(inBuf);
    P:=outBuf;
    limit:=P+outputSize;
    indexPrefix:=fetchLine(P);
    pricesUrl:=fetchLine(P);
    mainPricesPrefix:=fetchLine(P)+pricesUrl;
    regionalPricesPrefix:=fetchLine(P)+pricesUrl;
    constituentsUrl:=fetchLine(P);
    namePrefix:=fetchLine(P);
    mainConstituentsPrefix:=fetchLine(P)+constituentsUrl;
    indexSuffix:=fetchLine(P);
    constituentsSuffix:=fetchLine(P);
    constituentsCounter:=fetchLine(P);
    constituentsInc:=StrToInt(fetchLine(P));
    fundamentalsUrl:=fetchLine(P);
    nonUSsuffix:=fetchLine(P);
    lstMarkets.Clear;
    while P<limit do
      begin
      line:=fetchLine(P);
      if line='' then
        begin
        //new url section
        currentUrlCountry:=fetchLine(P);
        line:=fetchLine(P);
        end;
      if (line[1]<='9') and (line[1]>='0') then
        begin
        closingTime:=StrToInt(Copy(line,1,4));
        indexCountry:=Copy(line,5,255)+' ';
        line:=fetchLine(P);
        end;
      line:=line+',';
      i:=1;
      while line[i]<>',' do inc(i);
      mktName:=Copy(line,1,i-1);
      marketSymbols:=marketSymbols+mktName+StringOfChar(' ',nameLength-(i-1));
      inc(i);
      start:=i;
      while line[i]<>',' do inc(i);
      lstMarkets.Items.Add(indexCountry+Copy(line,start,i-start)+' ('+mktName+')');
      SetLength(urlCountry,lstMarkets.Items.Count);
      urlCountry[length(urlCountry)-1]:=currentUrlCountry;
      SetLength(closingTimes,length(urlCountry));
      closingTimes[length(urlCountry)-1]:=closingTime;
      end;
    FreeMem(outBuf);
  except marketSymbols:=''; end;
  end;
numMarkets:=0;//will be number of ticked markets
if marketSymbols='' then
  begin
  SetLength(urlCountry,1);
  SetLength(closingTimes,1);
  DeleteFile(marketsFileName);
  errorMessage('The ' + marketsFileName + ' file is missing. Please reinstall the software.');
  end;
problems:='';//a list of market files that could not be opened or were corrupted
if (FindFirst({ExtractFileDir(Application.ExeName)+'\}'*.mkt', faAnyFile, searchResult) = 0) then
  begin
  Repeat
    mktName:=Copy(searchResult.Name,1,length(searchResult.Name)-4);
    if length(mktName)<=nameLength then
      begin
      i:=Pos(padName(mktName),marketSymbols);
      if i=0 then
      //add imported shares/indices to list
        begin
        lstMarkets.Items.Add(mktName);
        marketSymbols:=marketSymbols+padName(mktName);
        i:=length(marketSymbols);
        end;
      if (searchResult.Attr and faReadOnly)=0 then
        begin
        allocMarketSpace(numMarkets+1);
        i:=(i-1) div nameLength;
        markets[numMarkets]:=i;
        if not getPricesFile(mktName,numMarkets,true) and not getPricesFile(mktName,numMarkets,false) then
          begin
          problems:=problems+mktName+', ';
          allocMarketSpace(numMarkets);//unallocate space
          end
        else
          begin
          lstMarkets.Checked[i]:=true;
          inc(numMarkets);
          end;
        end;
      end;
  Until FindNext(searchResult) <> 0;
  FindClose(searchResult);
  end;
if problems>'' then
  errorMessage('The price files for the following indices are corrupted: '
  +Copy(problems,1,length(problems)-2));
end;

function fetchLine(var P:PChar):String;
var
  start:PChar;
begin
start:=P;
while P^<>chr(13) do inc(P);
SetLength(Result,P-start);
Move(start^,Pointer(Result)^,P-start);
inc(P,2);
end;

procedure TForm1.lstMarketsClickCheck(Sender: TObject);
var
  mktName:String;
  i,j: Integer;
  response:Word;
begin
for i:=0 to lstMarkets.Items.Count-1 do
  begin
  mktName:=TrimRight(Copy(marketSymbols,i*nameLength + 1,nameLength));
  if (not (lstMarkets.Checked[i] xor ((FileGetAttr(mktName+'.mkt') and faReadOnly)>0)))
  or (not lstMarkets.Checked[i] and not FileExists(mktName+'.mkt')) then
    begin
    //state changed
    //determine which marketNum this item refers to, if any
    j:=mktArrayIndex(i);
    if lstMarkets.Checked[i] then
      begin
      //load in data from file
      if FileExists(mktName+'.mkt') or FileExists(mktName+'.bak') then
        begin
        allocMarketSpace(numMarkets+1);
        j:=numMarkets;
        markets[numMarkets]:=i;
        inc(numMarkets);
        if not getPricesFile(mktName,j,true) and not getPricesFile(mktName,j,false) then
          begin
          errorMessage('The price file and its backup for the '+mktName+ ' index are corrupted and will now be deleted. '
          + 'Please click the Update button to download the historical prices again.');
          DeleteFile(mktName+'.mkt');
          DeleteFile(mktName+'.bak');
          markets[j]:=0;
          end
        else
          begin
          //mark file as writeable
          FileSetAttr(mktName+'.mkt',0);
          //re-populate lists to include new index constituents
          refreshLists;
          end;
        end
      else if mktArrayIndex(i)<0 then //not already allocated
        begin
        inc(numMarkets);
        allocMarketSpace(numMarkets);
        markets[numMarkets-1]:=i;
        //allocate enough space for overall market price (don't know numCompanies yet)
        //if i>0 then SetLength(yahooPrices[numMarkets-1],1,numDays);
        end;
      end
    else if j>=0 then
      begin
      if pricesChanged[j] then
        begin
        response:=MessageDlg('The prices for ' + mktName
       +  ' have not been saved since being updated. Do you want to save them now, otherwise '
       + 'the changes will be lost?',mtConfirmation, [mbYes, mbNo, mbCancel], 0);
        if response=mrCancel then
          begin
          lstMarkets.Checked[i]:=true;
          Exit;
          end
        else if response=mrYes then
          saveMarket(mktName,j);
        end;
      //previously ticked, so zero arrays to release memory
      //or if this is last array, reduce numMarkets
      if j=numMarkets-1 then
        begin
        dec(numMarkets);
        allocMarketSpace(numMarkets);
        end
      else
        begin
        SetLength(namesArray[j],0);
        SetLength(yahooPrices[j],0);
        SetLength(datesArray[j],0);
        markets[j]:=-1;
        end;
      FileSetAttr(mktName+'.mkt',faReadOnly);//flag the file as read-only
      refreshLists;
      end;
    end;
  end;
btnSave.Enabled:=anyPricesChanged;
updatePortfolioGrid;
end;

function mktArrayIndex(i:Integer):Integer;
begin
Result:=numMarkets-1;
while (Result>=0) and (markets[Result]<>i) do
  dec(Result);
end;

procedure TForm1.lstSymbolChange(Sender: TObject);
begin
if portfolioChanged and (MessageDlgPos('Save changes to '''+lstPortfolio.Items[editedIndex]
+'''?',mtConfirmation, [mbYes, mbNo], 0, Left+261, Top+132) = mrYes) then
  savePortfolio;
prefix:=lstSymbol.ItemIndex=0;
if addSymbol then
  refreshLists
else
  updatePortfolioGrid;
end;

procedure TForm1.chkSymbolClick(Sender: TObject);
begin
if hintDown(Sender) then Exit;
addSymbol:=chkSymbol.Checked;
refreshLists;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
mHHelp.Free;
//HHCloseAll;
end;

procedure TForm1.FormClick(Sender: TObject);
begin
hintDown(Sender);
end;

procedure TForm1.lstMarketsClick(Sender: TObject);
begin
if hintDown(Sender) then Exit;
if not lstMarkets.Checked[lstMarkets.ItemIndex] then
  lstMarkets.ItemIndex:=-1
else
  updateRightControls(lstMarkets.Items[lstMarkets.ItemIndex]);  
btnExport.Enabled:=lstMarkets.ItemIndex>=0;
btnClear.Enabled:=btnExport.Enabled;
end;

procedure TForm1.chkSpreadBetClick(Sender: TObject);
var
  state: Boolean;
begin
if hintDown(Sender) then
  chkSpreadBet.Checked:=not chkSpreadBet.Checked
else
  begin
  state:=chkSpreadBet.Checked;
  lblMinBet.Enabled:=state;
  txtMinBet.Enabled:=state;
  lblMaxBet.Enabled:=state;
  spnMaxBet.Enabled:=state;
  lblIMR.Enabled:=state;
  spnMargin.Enabled:=state;
  lblSpread.Enabled:=state;
  spnSpread.Enabled:=state;
  end;
end;

procedure TForm1.chkClearResultsClick(Sender: TObject);
begin
if hintDown(Sender) then chkClearResults.Checked:=not chkClearResults.Checked;
end;

function strToPrice(const priceStr:String):Single;
var
  code:Integer;
begin
Val(priceStr,Result,code);
if code>0 then Result:=0;
end;

procedure TForm1.btnImportClick(Sender: TObject);
var
  F:TextFile;
  S,mktName,symbol,fileName,mktSuffix:String;
  i,j,c,market,dateNo,shareNo,selectedIndex:Integer;
  companyLen,lineCount,len,noDays,dateOffset,pricePoint:Integer;
  year,month,day,dateWord:Word;
  newMarket,tomorrowCheck,googleFile:Boolean;
  mktAvg,price:Single;
label
  formatError;
begin
if hintDown(Sender) then Exit;
if orderId=0 then
  TfrmNag.Create(self).ShowModal
else if OpenDialog1.Execute then
  begin
  Application.ProcessMessages;
  Screen.Cursor:=crHourglass;
  selectedIndex:=lstMarkets.ItemIndex;
  if selectedIndex>0 then
    begin
    mktName:=getSymbol(selectedIndex);
    market:=mktArrayIndex(selectedIndex);
    end;
  for c:=0 to OpenDialog1.Files.Count-1 do
    begin
    fileName:=ExtractFileName(OpenDialog1.Files.Strings[c]);
    i:=Pos('_',fileName);
    if (selectedIndex<=0) or (i=0) then
      begin
      if i=0 then
        mktName:=LStr(fileName,ScanB(fileName,'.',0)-1)
      else
        mktName:=Copy(fileName,1,i-1);
      j:=lstMarkets.Items.IndexOf(mktName);
      if j>=0 then
        begin
        FileSetAttr(mktName+'.mkt',0);
        market:=mktArrayIndex(j);
        end
      else
        begin
        j:=lstMarkets.Items.Add(mktName);
        marketSymbols:=marketSymbols+padName(mktName);
        market:=-1;
        end;
      if market<0 then
        begin
        allocMarketSpace(numMarkets+1);
        market:=numMarkets;
        markets[market]:=j;
        inc(numMarkets);
        end;
      lstMarkets.Checked[j]:=true;
      end;
    if i=0 then //not an eod file
      begin
      with TFileStream.Create(OpenDialog1.Files.Strings[c],fmOpenRead or fmShareDenyNone) do
        begin
        SetLength(S,Size);
        Read(Pointer(S)^,Size);
        Free;
        end;
      namesArray[market]:=padName(mktName);
      processPrices(S,market);
      if not pricesChanged[market] then
        begin
        errorMessage('The file named '+fileName+' is not recognized as a valid price file.');
        Screen.Cursor:=crDefault;
        Exit;
        end;
      end
    else
      begin
      if market<0 then
        begin
        errorMessage('The '+TrimRight(mktName)+' index must be ticked before price files can be imported into it.');
        Break;
        end;
      year:=StrToInt(Copy(fileName,i+1,4));
      month:=StrToInt(Copy(fileName,i+5,2));
      day:=StrToInt(Copy(fileName,i+7,2));
      dateWord:=dateToWord(day,month,year);
      noDays:=length(datesArray[market]);
      dateNo:=noDays-1;
      while (dateNo>=0) and (datesArray[market,dateNo]>dateWord) do dec(dateNo);
      newMarket:=length(namesArray[market])=0;
      if (dateNo<0) or (datesArray[market,dateNo]<>dateWord) then //date doesn't already exist
        begin
        inc(dateNo);
        dateOffset:=noDays-dateNo;
        inc(noDays);
        SetLength(datesArray[market],noDays);
        Move(datesArray[market,dateNo],datesArray[market,dateNo+1],2*dateOffset);
        datesArray[market,dateNo]:=dateWord;
        for shareNo:=1 to numCompanies(market)-1 do
          begin
          companyLen:=length(yahooPrices[market,shareNo]);
          SetLength(yahooPrices[market,shareNo],companyLen+1);
          pricePoint:=companyLen-dateOffset;
          if pricePoint<0 then pricePoint:=0;
          Move(yahooPrices[market,shareNo,pricePoint],yahooPrices[market,shareNo,pricePoint+1],4*(companyLen-pricePoint));
          yahooPrices[market,shareNo,pricePoint]:=0;
          end;
        end
      else
        dateOffset:=noDays-dateNo-1;
      AssignFile(F,OpenDialog1.Files.Strings[c]);
      Reset(F);
      if newMarket then
        begin
        i:=-1;
        Repeat
          Readln(F);
          inc(i);
        Until Eof(F);
        SetLength(yahooPrices[market],i);
        for shareNo:=1 to i-1 do
          begin
          SetLength(yahooPrices[market,shareNo],1);
          yahooPrices[market,shareNo,0]:=0;
          end;
        Reset(F);
        namesArray[market]:=padName('^'+mktName);
        end;
      Readln(F,S);
      if ScanF(S,'date',-1)>0 then Readln(F,S);
      lineCount:=0;
      //determine if existing shares have an index suffix
      mktSuffix:='';
      if length(namesArray[market])>nameLength then //more than just index
        begin
        i:=2*nameLength-1;
        Repeat dec(i) Until (namesArray[market][i]='.') or (i=nameLength);
        if i>nameLength then //existing shares have suffix
          begin
          mktSuffix:=RTrim(Copy(namesArray[market],i,2*nameLength-i),' ');
          i:=numCompanies(market);
          j:=length(mktSuffix);
          Repeat
            dec(i);
          Until (i=0) or (RStr(RTrim(getName(market,i),' '),j)<>mktSuffix);
          if i>0 then
            mktSuffix:='' //don't add suffix if not all existing shares have it
          else
            begin
            i:=2;
            while S[i]<>',' do inc(i);
            j:=i;
            Repeat dec(i) Until (S[i]='.') or (i=2);
            if S[i]=symbolSeparator then
              begin
              if Copy(S[i],i+1,j-i-1)<>mktSuffix then
                begin
                errorMessage('The file '+fileName+' cannot be imported into index '
                +mktName+' as it belongs to a different market.');
                Break;
                end
              else
                mktSuffix:='';
              end;
            end;
          end;
        end;
      mktAvg:=0;
      while not Eof(F) do
        begin
        //read symbol
        i:=1;
        while S[i]<>',' do inc(i);
        symbol:=padName(Copy(S,1,i-1)+mktSuffix);
        if not newMarket then shareNo:=Pos(symbol,namesArray[market]);
        if newMarket or (shareNo=0) then
          begin
          if selectedIndex<0 then //don't add shares to index from a different source
            begin
            //new share
            shareNo:=numCompanies(market);
            namesArray[market]:=namesArray[market]+symbol;
            pricePoint:=0;
            if not newMarket then
              begin
              SetLength(yahooPrices[market],shareNo+1);//first share is index price
              SetLength(yahooPrices[market,shareNo],dateOffset+1);
              end;
            end
          else
            shareNo:=0;
          end
        else
          begin
          shareNo:=(shareNo-1) div nameLength;
          pricePoint:=length(yahooPrices[market,shareNo])-dateOffset-1;
          end;
        if (shareNo>0) and (pricePoint>=0) then
          begin
          i:=length(S);
          while S[i]<>',' do dec(i);
          dec(i);
          j:=i;
          while S[i]<>',' do dec(i);
          price:=strToPrice(Copy(S,i+1,j-i));
          yahooPrices[market,shareNo,pricePoint]:=price;
          mktAvg:=mktAvg+price;
          end;
        Readln(F,S);
        inc(lineCount);
        end;
{      tomorrowCheck:=true;//double as "successful import" flag
      //check for holiday
      if not newMarket then
        Repeat
        tomorrowCheck:=tomorrowCheck xor true;
        if (not tomorrowCheck and (dateNo>0)) or (tomorrowCheck and (dateNo<noDays-1)) then
          begin
          if tomorrowCheck then dateToCheck:=dateNo+1 else dateToCheck:=dateNo-1;
          j:=0;
          for i:=0 to oldNoCompanies-1 do
            if yahooPrices[market,i,dateNo]=yahooPrices[market,i,dateToCheck] then inc(j);
          if j > 3 * oldNoCompanies div 4 then
            begin
            if not tomorrowCheck then inc (dateToCheck);
            if MessageDlg('The day ' + dayToStr(market,dateToCheck) + ' seems to be'
            +' a holiday for the ' + TrimRight(mktName) +' market. Do you want to keep it?'
            , mtConfirmation, [mbYes, mbNo], 0)=mrNo then
              begin
              if tomorrowCheck then
                begin
                datesArray[market,dateNo+1]:=datesArray[market,dateNo];//tomorrow is holiday
                pricesChanged[market]:=true;//make sure date changed is savable
                end;
              Move(datesArray[market,0],datesArray[market,1],2*dateNo);
              datesArray[market,0]:=lostDate;
              for i:=0 to oldNoCompanies-1 do
                begin
                Move(yahooPrices[market,i,0],yahooPrices[market,i,1],4*length(yahooPrices[market,i]));
                yahooPrices[market,i,0]:=lostPrices[i];
                end;
              i:=length(lostPrices)*nameLength;
              Delete(namesArray[market],i,length(namesArray[market])-i);
              tomorrowCheck:=false;
              Break;
              end;
            end;
          end;
        Until tomorrowCheck;
      if tomorrowCheck then //successful import
        begin }
        SetLength(yahooPrices[market,0],noDays);
        Move(yahooPrices[market,0,dateNo],yahooPrices[market,0,dateNo+1],4*dateOffset);
        yahooPrices[market,0,dateNo]:=mktAvg/lineCount;
        pricesChanged[market]:=true;
//        end;
//      end;
      CloseFile(F);
      end;
    end;
  for market:=0 to numMarkets-1 do
    if pricesChanged[market] then processDelistedShares(market);
  Screen.Cursor:=crDefault;
  refreshLists;
  end;
end;

procedure processPrices(const S:String; market:Integer);
var
  i,j,k,dateNo:Integer;
  commaCount,len:Integer;
  year,month,day:Word;
  googleFile:Boolean;
begin
dateNo:=CountF(S,chr(10),1)-1;
SetLength(datesArray[market],dateNo);
SetLength(yahooPrices[market],1,dateNo);
if Pos('Close',S)>0 then //yahoo or google file
  begin
  googleFile:=(Pos('Adj Close',S)=0);
  if Form1.chkUnadjusted.Checked or googleFile then commaCount:=3 else commaCount:=5;
  dec(dateNo);
  i:=1;
  while S[i]<>chr(10) do inc(i);//skip header
  inc(i);
  len:=length(S);
  //populate dates array (date is 1st column of each line)
  //and overall index price (closing is 5th column)
  while (dateNo>=0) and (i<len) do
    begin
    //note: hist date format is yyyy-mm-dd
    try
      if googleFile then
        begin
        if S[i+1]='-' then
          begin
          day:=StrToInt(S[i]);
          inc(i,2);
          end
        else
          begin
          day:=StrToInt(S[i]+S[i+1]);
          inc(i,3);
          end;
        month:=(Pos(Copy(S,i,3),months)+2) div 3;
        year:=StrToInt(Copy(S,i+4,2));
        if year>28 then inc(year,1900) else inc(year,2000);
        inc(i,7);
        end
      else
        begin
        year:=StrToInt(Copy(S,i,4));
        month:=StrToInt(Copy(S,i+5,2));
        day:=StrToInt(Copy(S,i+8,2));
        inc(i,11);
        end;
      datesArray[market,dateNo]:=dateToWord(day,month,year);
    except
      Exit;
    end;
    //move to start of Close or Adj. Close
    j:=commaCount;
    Repeat
      while (S[i]<>',') do inc(i);
      inc(i);
      dec(j);
    Until j=0;
    j:=i;
    //move to next comma or end of line
    if commaCount=3 then
      begin
      Repeat inc(i) Until S[i]=',';
      yahooPrices[market,0,dateNo]:=strToPrice(Copy(S,j,i-j));
      while S[i]<>chr(10) do inc(i);//move to next line
      end
    else
      begin
      Repeat inc(i) Until S[i]=chr(10);
      yahooPrices[market,0,dateNo]:=strToPrice(Copy(S,j,i-j));
      end;
    dec(dateNo);
    inc(i);
    end;
  end
else //assume two column file: date, price
  begin
  j:=1;
  for i:=0 to dateNo-1 do
    begin
    try
      day:=StrToInt(Copy(S,j,2));
      month:=StrToInt(Copy(S,j+3,2));
      year:=StrToInt(Copy(S,j+6,4));
      //if month=0 then goto formatError;
      datesArray[market,i]:=day + (month shl 5) + ((year mod 100) shl 9);
    except
      Exit;
      end;
    inc(j,11);
    k:=j;
    while S[j]<>chr(10) do inc(j);
    if S[j-1]=chr(13) then
      yahooPrices[market,0,i]:=strToPrice(Copy(S,k,j-k-1))
    else
      yahooPrices[market,0,i]:=strToPrice(Copy(S,k,j-k));
    inc(j);
    end;
  end;
pricesChanged[market]:=true;
end;

procedure processDelistedShares(market:Integer);
var
  PNames:PChar;
  cName:String;
  i,noCompanies,len:Integer;
  tempPricesArray:array of Single;
begin
PNames:=PChar(namesArray[market]);
//reinstate any so-called delisted shares that are now non-zero
//note: [0] is always index itself
//noDays:=length(datesArray[market]);
i:=1;
while i<=numDelisted[market] do
  begin
//for i:=1 to numDelisted[market] do
  if yahooPrices[market,i,length(yahooPrices[market,i])-1]>0 then
    begin
    //swap share with one at end of delisted block
    //swap names
    cName:=getName(market,i);
    Move((PNames+numDelisted[market]*nameLength)^,(PNames+i*nameLength)^,nameLength);
    Move(PChar(cName)^,(PNames+numDelisted[market]*nameLength)^,nameLength);
    //swap prices
    SetLength(tempPricesArray,length(yahooPrices[market,i]));
    Move(yahooPrices[market,i,0],tempPricesArray[0],4*length(tempPricesArray));
    SetLength(yahooPrices[market,i],length(yahooPrices[market,numDelisted[market]]));
    Move(yahooPrices[market,numDelisted[market],0],yahooPrices[market,i,0],4*length(yahooPrices[market,i]));
    SetLength(yahooPrices[market,numDelisted[market]],length(tempPricesArray));
    Move(tempPricesArray[0],yahooPrices[market,numDelisted[market],0],4*length(tempPricesArray));
    dec(numDelisted[market]);
    end
  else
    inc(i);
  end;
//determine delisted shares
i:=numDelisted[market]+1;
noCompanies:=numCompanies(market);
while i<noCompanies do
  begin
  len:=length(yahooPrices[market,i]);
  if len=0 then //remove blank share
    begin
    //overwrite share with last one in index
    dec(noCompanies);
    cName:=getName(market,i);
    Move((PNames+noCompanies*nameLength)^,(PNames+i*nameLength)^,nameLength);
    Delete(namesArray[market],noCompanies*nameLength+1,nameLength);
    SetLength(yahooPrices[market,i],length(yahooPrices[market,noCompanies]));
    Move(yahooPrices[market,noCompanies,0],yahooPrices[market,i,0],4*length(yahooPrices[market,i]));
    SetLength(yahooPrices[market],noCompanies);
    end
  else if yahooPrices[market,i,len-1]<=0 then
    begin
    //swap delisted share with the one just after end of delisted group
    cName:=getName(market,i);
    //swap names
    inc(numDelisted[market]);//also moves to one after index
    Move((PNames+numDelisted[market]*nameLength)^,(PNames+i*nameLength)^,nameLength);
    Move(PChar(cName)^,(PNames+numDelisted[market]*nameLength)^,nameLength);
    //swap prices
    SetLength(tempPricesArray,length(yahooPrices[market,i]));
    Move(yahooPrices[market,i,0],tempPricesArray[0],4*length(tempPricesArray));
    SetLength(yahooPrices[market,i],length(yahooPrices[market,numDelisted[market]]));
    Move(yahooPrices[market,numDelisted[market],0],yahooPrices[market,i,0],4*length(yahooPrices[market,i]));
    SetLength(yahooPrices[market,numDelisted[market]],length(tempPricesArray));
    Move(tempPricesArray[0],yahooPrices[market,numDelisted[market],0],4*length(tempPricesArray));
    end;
  inc(i);
  end;
end;

procedure TForm1.ShowFundamentals1Click(Sender: TObject);
begin
ShellExecute(Handle, 'open', PChar('http://'+urlCountry[markets[currMarket]]
+fundamentalsUrl+getName(currMarket,companyNo)),nil,nil, SW_SHOWNORMAL)
end;

procedure TForm1.btnExportClick(Sender: TObject);
var
  F: TextFile;
  i,company,market,start,finish,noCompanies:Integer;
  S,partNoStr,baseName:String;
  p:Single;
begin
if hintDown(Sender) then Exit;
if orderId=0 then
  begin
  TfrmNag.Create(self).ShowModal;
  Exit;
  end;
baseName:=lstMarkets.Items[lstMarkets.ItemIndex];
i:=Pos('(',baseName);
if i>0 then baseName:=Copy(baseName,i+1,length(baseName)-i-1);
market:=mktArrayIndex(lstMarkets.ItemIndex);
noCompanies:=numCompanies(market)-1;
if noCompanies<256 then
  partNoStr:=''
else
  partNoStr:=' (1)';
with SaveDialog1 do
  begin
  Title:='Export prices';
  FileName:=baseName+partNoStr;
  OpenDialog1.Filter := '*.csv|*.csv';
  if not Execute then Exit;
  end;
Application.ProcessMessages;
Screen.Cursor:=crHourGlass;
try
  start:=0;
  while start<noCompanies do
    begin
    AssignFile(F,baseName+partNoStr+'.csv');
    Rewrite(F);
    finish:=start+254;
    if finish>=noCompanies then finish:=noCompanies-1;
    S:='Date';
    for company:=start to finish do
      S:=S+','+companyName(market,company);
    WriteLn(F,S);
    for i:=0 to length(datesArray[market])-1 do
      begin
      S:=dayToStr(market,i);
      for company:=start to finish do
        begin
        p:=Abs(yahooPrices[market,company,i]);
        if Frac(p)=0 then
          S:=S+','+IntToStr(Trunc(p))
        else
          S:=S+','+FloatToStrF(p,ffFixed,7,2);
        end;
      WriteLn(F,S);
      end;
    CloseFile(F);
    inc(start,255);
    partNoStr:=' ('+IntToStr(1 + start div 255)+')';
    end;
except
  errorMessage('Could not export prices. Please check that the target file is not already open.');
  end;
Screen.Cursor:=crDefault;
end;

procedure TForm1.grdPriceGetEditText(Sender: TObject; ACol, ARow: Integer;
  var Value: String);
begin
editPrice:=Value;
priceGridDay:=ACol;
end;

procedure TForm1.spnMaxBetClick(Sender: TObject);
begin
hintDown(Sender);
end;

procedure TForm1.spnTestAv1Click(Sender: TObject);
begin
hintDown(Sender);
end;

procedure TForm1.spnAverage1Click(Sender: TObject);
begin
hintDown(Sender);
end;

procedure TForm1.optTestExpMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
hintDown(Sender);
end;

procedure TForm1.lstPortfolioExit(Sender: TObject);
begin
if lstPortfolio.ItemIndex<lstPortfolio.Items.Count-1 then
  editPortfolioName;
end;

procedure TForm1.grdPortfolioEnter(Sender: TObject);
begin
AddToPortfolio.Visible:=false;
end;

procedure TForm1.printGraphClick(Sender: TObject);
begin
if not PrintDialog1.Execute then Exit;
Application.ProcessMessages;//close print dialog immediately
with Printer do
  begin
 // Orientation:=poLandscape;
  Title:=selectedName+' graph';
  BeginDoc;
  with Canvas do
    begin
    Font.Name:='Times New Roman';
    Font.Size:=14;
    Pen.Color:=clBlack;
    TextOut((PageWidth-TextWidth(selectedName)) div 2, 100, selectedName);
    Font.Size:=8;
    Pen.Width:=5;
    end;
  plotGraph(false);
  EndDoc;
  end;
end;

procedure TForm1.saveGraphClick(Sender: TObject);
var
  S, ext: String;
  bitmap: TBitmap;
  i: Integer;
begin
with SaveDialog1 do
  begin
  S:=selectedName;
  for i:=1 to length(S) do
    if S[i]=symbolSeparator then S[i]:='_';
  FileName:=S;
  Title:='Save graph to file';
  Filter := 'Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg';
  if not Execute then Exit;
  S:=FileName;
  end;
Application.ProcessMessages;//to ensure dialog does not obscure graph
bitmap:=TBitmap.Create;
graphToBitmap(bitmap);
ext:=ExtractFileExt(S);
if ext='.bmp' then
  bitmap.SaveToFile(S)
else if ext='.jpg' then with TJpegImage.Create do
  begin
  Assign(bitmap);
  SaveToFile(S);
  Free;
  end
else
  errorMessage('Could not save graph. Please use an extension of either bmp or jpg.');
bitmap.Free;
end;

procedure TForm1.graphToBitmap(out bitmap:TBitmap);
var
  source, dest: TRect;
begin
with bitmap do
  begin
  Width:=Graph.Width;
  Height:=Graph.Height;
  dest:=Rect(0, 0, Width, Height);
  source:=Rect(0, 0, Graph.Width, Graph.Height);
  Canvas.CopyRect(dest, Graph.Canvas, source);
  end;
end;

procedure TForm1.copyGraphClick(Sender: TObject);
var
  bitmap:TBitmap;
begin
Application.ProcessMessages;
bitmap:=TBitmap.Create;
graphToBitmap(bitmap);
Clipboard.Assign(bitmap);
bitmap.Free;
end;

procedure TForm1.edtSearchChange(Sender: TObject);
var
  searchText: String;
  score,bestScore,i,best:Integer;
begin
if (sharesListed) then
  begin
  searchText:=UpperCase(edtSearch.Text);
  //search the names string for any occurrences of exact search string
  bestScore:=0;
  best:=0;
  for i:=0 to lstAllShares.Items.Count-1 do
    begin
    score:=Similar(searchText,UpperCase(lstAllShares.Items[i]));
    if (score>bestScore) and (score<=100) then
      begin
      bestScore:=score;
      best:=i;
      end;
    end;
  if bestScore>0 then
    begin
    lstAllShares.ItemIndex:=best;
    lstAllSharesClick(Sender);
    end;
  end;
end;

procedure TForm1.grdPortfolioMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
if portfolioMouseDown then
  begin
  grdPortfolio.ColWidths[0]:=X;
  grdPortfolio.Refresh;
  end
else if (Y<28) and (Abs(X-grdPortfolio.ColWidths[0])<5) then
  grdPortfolio.Cursor:=crHSplit
else
  grdPortfolio.Cursor:=crDefault;
end;

procedure TForm1.grdPortfolioSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
var
  dateStr,symbol,newName:String;
  newMarket,newCompany:Integer;
  dateWord:Word;
begin
with grdPortfolio do begin
if ARow=RowCount-1 then
  btnRemove.Enabled:=false
else
  begin
  btnRemove.Enabled:=true;
  if Cells[portfolioCol[10],ARow]<>defaultStr then
    begin
    simpleAvg:=extractAvgSettings(Cells[portfolioCol[10],ARow],lowAverage,highAverage);
    lblAvg1.Show;
    if (lowAverage>1) then
      begin
      lblAvg1.Caption:=IntToStr(lowAverage)+'-day';
      lblAvg2.Show;
      lblAvg2.Caption:=IntToStr(highAverage)+'-day';
      end
    else
      begin
      lblAvg2.Hide;
      lblAvg1.Caption:=IntToStr(highAverage)+'-day';
      end;
    end
  else
    changeSpinValues(true);
  if addSymbol then
    updateRightControls(Cells[0,ARow])
  else
    begin
    extractNameAndSymbol(Cells[0,ARow],newName,symbol);
    newMarket:=mktArrayIndex((Pos(symbol,marketSymbols)-1) div nameLength);
    if newMarket>=0 then
      begin
      newCompany:=Pos(newName,namesArray[newMarket])-1;
      if newCompany>=0 then
        begin
        companyNo:=newCompany div nameLength;
        currMarket:=newMarket;
        selectedName:=TrimRight(newName);
        lstAllShares.ItemIndex:=lstAllShares.Items.IndexOf(selectedName);
        refreshGraphSelection;
        populatePriceGrid;
        ShowFundamentals1.Visible:=(currMarket>=0) and (markets[currMarket]>0) and (markets[currMarket]<length(urlCountry));
        end;
      end;
    end;
  if portfolioCol[6]=ACol then //purchase date column
    begin
    editRow:=ARow;
    dateStr:=Cells[ACol,ARow];
    if dateStr<>'-' then
      begin
      dateWord:=dateStrToWord(dateStr);
      DateTimePicker1.Date:=EncodeDate(1950+(dateWord shr 9),(dateWord shr 5) and 15,dateWord and 31);
      end
    else
      DateTimePicker1.Date:=Date;
    positionDatePicker;
    DateTimePicker1.Top:=(ARow+2-TopRow)*(DefaultRowHeight+1)+Top-5;
    DateTimePicker1.Show;
    end
  else
    DateTimePicker1.Hide;
  end;
end;
if not DateTimePicker1.Visible then changePortfolioPrice;
PopUpMenu1.Autopopup:=btnRemove.Enabled;
end;

procedure TForm1.deleteDayClick(Sender: TObject);
var
  noDays,i,dateOffset,companyLen,dateNo:Integer;
begin
noDays:=length(datesArray[currMarket]);
if noDays=0 then exit;
dateNo:=grdPrice.Col;
if dateNo<noDays-1 then
  Move(datesArray[currMarket,dateNo+1],datesArray[currMarket,dateNo],2*(noDays-dateNo-1));
for i:=0 to numCompanies(currMarket)-1 do
  begin
  companyLen:=length(yahooPrices[currMarket,i]);
  if companyLen>dateNo then
    begin
    Move(yahooPrices[currMarket,i,dateNo+1],yahooPrices[currMarket,i,dateNo],4*(companyLen-dateNo-1));
    SetLength(yahooPrices[currMarket,i],companyLen-1);
    end;
  end;
SetLength(datesArray[currMarket],noDays-1);
populatePriceGrid;
pricesChanged[currMarket]:=true;
refreshGraphSelection;
refreshGraph;
populateUpDownLists;
updateMoversLists;
updateHighLowLists;
btnSave.Enabled:=True;
end;

procedure TForm1.Addday1Click(Sender: TObject);
var
  year,month,day:Word;
  S:String;
begin
S:=InputBox('Insert price', 'Date for new price:', '');
if S='' then exit;
DecodeDate(StrToDate(S),year,month,day);
if day=0 then exit;
insertDay(day,month,year);
end;

procedure insertDay(day,month,year:Word);
var
  dateWord:Word;
  noDays,i,j,dateOffset,companyLen,dateNo:Integer;
begin
dateWord:=dateToWord(day,month,year);
noDays:=length(datesArray[currMarket]);
i:=noDays-1;
while datesArray[currMarket,i]>dateWord do dec(i);
if datesArray[currMarket,i]=dateWord then
  errorMessage('This date already exists.')
else
  begin
  SetLength(datesArray[currMarket],1+noDays);
  inc(i);
  dateOffset:=noDays-i;
  Move(datesArray[currMarket,i],datesArray[currMarket,i+1],2*dateOffset);
  datesArray[currMarket,i]:=dateWord;
  for j:=0 to numCompanies(currMarket)-1 do
    begin
    companyLen:=length(yahooPrices[currMarket,j]);
    SetLength(yahooPrices[currMarket,j],companyLen+1);
    dateNo:=companyLen-dateOffset;
    Move(yahooPrices[currMarket,j,dateNo],yahooPrices[currMarket,j,dateNo+1],4*dateOffset);
    yahooPrices[currMarket,j,dateNo]:=0;
    end;
  end;
Form1.populatePriceGrid;
end;

procedure TForm1.graphScrollChange(Sender: TObject);
begin
if sharesListed then
  begin
  if lstRange.ItemIndex<7 then
    begin
    selectionEnd:=graphScroll.Position;
    selectionStart:=Max(0,selectionEnd-rangeOptions[lstRange.ItemIndex]);
    showPrice.Enabled:=(selectionEnd=length(yahooPrices[currMarket,companyNo])-1);
    end
  else
    begin
    selectionStart:=0;
    selectionEnd:=length(yahooPrices[currMarket,companyNo])-1;
    showPrice.Enabled:=true;
    end;
  refreshGraph;
  end;
end;

procedure TForm1.chkListTradesClick(Sender: TObject);
begin
if hintDown(Sender) then chkListTrades.Checked:=not chkListTrades.Checked;
end;

procedure TForm1.chkOptimizeClick(Sender: TObject);
begin
if hintDown(Sender) then chkOptimize.Checked:=not chkOptimize.Checked;
end;

procedure TForm1.Editday1Click(Sender: TObject);
var
  year,month,day:Word;
  S:String;
begin
S:=InputBox('Edit date', 'New date:', '');
if S='' then exit;
DecodeDate(StrToDate(S),year,month,day);
if day>0 then
  begin
  datesArray[currMarket,priceGridDay]:=dateToWord(day,month,year);
  populatePriceGrid;
  end;
end;

procedure TForm1.showPriceClick(Sender: TObject);
var
  price: Single;
begin
if simpleAvg then
  begin
  if (lowAverage>1) and (lowAverage<highAverage) then //two averages crossing
    price:=(highAverage*graphData[selectionEnd+1-lowAverage,0]-lowAverage*graphData[selectionEnd+1-highAverage,0]
    +lowAverage*highAverage*(graphData[selectionEnd,2]-graphData[selectionEnd,1]))/(highAverage-lowAverage)
  else if (highAverage>1) then //price crossing one average
    price:=(graphData[selectionEnd,1]*highAverage-graphData[selectionEnd+1-highAverage,0])/(highAverage-1);
  end
else
  begin
  if (lowAverage>1) and (lowAverage<highAverage) then
    begin //two averages crossing
    price:=(graphData[selectionEnd,2]*(highAverage-1)*(lowAverage+1)
    -graphData[selectionEnd,1]*(lowAverage-1)*(highAverage+1))/(2*(highAverage-lowAverage));
    end
  else if (highAverage>1) then //price crossing one average
    price:=graphData[selectionEnd,1];
  end;
ShowMessage('Price threshold for crossover the next day is '+priceToStr(price));
end;

procedure TForm1.lstTradeTypeChange(Sender: TObject);
begin
if btnGetResults.Caption='&Cancel' then cancelBacktest;
end;

procedure TForm1.Inserttoday1Click(Sender: TObject);
var
  year,month,day:Word;
begin
DecodeDate(Date,year,month,day);
insertDay(day,month,year);
end;

procedure TForm1.grdPortfolioContextPopup(Sender: TObject;
  MousePos: TPoint; var Handled: Boolean);
begin
resetAvgValues.Visible:=(grdPortfolio.Col=portfolioCol[10]);
end;

procedure TForm1.resetAvgValuesClick(Sender: TObject);
begin
changeSpinValues(false);
with grdPortfolio do Cells[Col,Row]:=avgValuesString;
portfolioAltered;
end;

end.


