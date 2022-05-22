unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, JvExComCtrls, JvComCtrls, JvExControls, JvLabel, StdCtrls,
  JvExStdCtrls, JvRichEdit, Grids, JvExGrids, JvStringGrid, JvStatusBar,
  ExtCtrls, Buttons, ImgList, JvAppStorage, JvAppIniStorage, JvComponentBase,
  JvFormPlacement, JvExExtCtrls, JvExtComponent, JvPanel, Menus, JvMenus;

type
  TMainForm = class(TForm)
    PageControl: TJvPageControl;
    OriginalTextSheet: TTabSheet;
    HeaderTextEdit: TJvRichEdit;
    JvLabel1: TJvLabel;
    JvLabel2: TJvLabel;
    Text2LoopEdit: TJvRichEdit;
    JvLabel3: TJvLabel;
    FooterTextEdit: TJvRichEdit;
    SRSheet: TTabSheet;
    FinalTextSheet: TTabSheet;
    JvLabel4: TJvLabel;
    SRGrid: TJvStringGrid;
    HelpSheet: TTabSheet;
    HelpTextEdit: TJvRichEdit;
    StatusBar: TJvStatusBar;
    SRAddBtn: TButton;
    SRDelBtn: TButton;
    FinalTextEdit: TJvRichEdit;
    JvLabel5: TJvLabel;
    ProcessTextByLoopBtn: TButton;
    CopyClpBtn: TButton;
    Save2FileBtn: TButton;
    SaveTextDialog: TSaveDialog;
    ImageList: TImageList;
    OpenPrjDialog: TOpenDialog;
    SavePrjDialog: TSaveDialog;
    FormStorage: TJvFormStorage;
    PrjFileStorage: TJvAppIniFileStorage;
    StatusTimer: TTimer;
    SRCopyBtn: TButton;
    SRPasteBtn: TButton;
    SRExcelExportBtn: TButton;
    SRExcelImportBtn: TButton;
    ProcessTextNormallyBtn: TButton;
    LogoPanel: TPanel;
    Logo: TImage;
    TopPanel: TJvPanel;
    NewPrjBtn: TButton;
    LoadPrjBtn: TButton;
    SavePrjBtn: TButton;
    Button1: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure SRAddBtnClick(Sender: TObject);
    procedure SRDelBtnClick(Sender: TObject);
    procedure ProcessTextByLoopBtnClick(Sender: TObject);
    procedure CopyClpBtnClick(Sender: TObject);
    procedure Save2FileBtnClick(Sender: TObject);
    procedure LoadPrjBtnClick(Sender: TObject);
    procedure SavePrjBtnClick(Sender: TObject);
    procedure StatusTimerTimer(Sender: TObject);
    procedure NewPrjBtnClick(Sender: TObject);
    procedure SRGridDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
      State: TGridDrawState);
    procedure SRCopyBtnClick(Sender: TObject);
    procedure SRPasteBtnClick(Sender: TObject);
    procedure SRExcelExportBtnClick(Sender: TObject);
    procedure SRExcelImportBtnClick(Sender: TObject);
    procedure ProcessTextNormallyBtnClick(Sender: TObject);
    procedure ProjectSheetContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
    procedure LogoClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure HelpTextEditChange(Sender: TObject);
  private
    { Private declarations }
    SRCopyCol: Integer;
    procedure ShowStatus(Text: string; Temporary: boolean = True);
    procedure SRGridInit;
    procedure ShowStatusPrj(Text: string);
  public
    { Public declarations }
  end;



  procedure AutoSizeCol(Grid: TStringGrid; Column: integer);
  procedure AutoSizeAllCol(Grid: TStringGrid);



var
  MainForm: TMainForm;
  PrjName: string;
  PrjPath: string;

implementation

{$R *.dfm}

uses ShellAPI, JclFileUtils, ExcelImportExport;



//
// Status info
//
procedure TMainForm.ShowStatus(Text: string; Temporary: boolean = True);
begin
  StatusTimer.Enabled := False;
  StatusBar.Panels[0].Text := Text;
  Application.ProcessMessages;
  StatusTimer.Enabled := Temporary;
end;


procedure TMainForm.ShowStatusPrj(Text: string);
begin
  StatusBar.Panels[1].Text := Text;
  Application.ProcessMessages;
end;


procedure TMainForm.StatusTimerTimer(Sender: TObject);
begin
  StatusBar.Panels[0].Text := '';
end;



//
// Form create/destroy
//
procedure TMainForm.FormCreate(Sender: TObject);
var
  Version: string;
begin
  OpenPrjDialog.InitialDir := ExtractFilePath(Application.ExeName)+'PROJECT';
  SavePrjDialog.InitialDir := OpenPrjDialog.InitialDir;
  SaveTextDialog.InitialDir := OpenPrjDialog.InitialDir;

  PageControl.TabIndex := 0;

  PrjName := 'NoName';

  Version := 'Version: '+VersionFixedFileInfoString(Application.ExeName, vfFull);
  ShowStatusPrj(Version);
  HelpTextEdit.Lines.Append(Version);

  SRGridInit;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
//
end;

procedure TMainForm.FormResize(Sender: TObject);
begin
  StatusBar.Panels[0].Width := Round(StatusBar.Width * 0.8);
  StatusBar.Panels[1].Width := Round(StatusBar.Width * 0.2);

//  AutoSizeAllCol(SRGrid);
end;






procedure TMainForm.HelpTextEditChange(Sender: TObject);
begin

end;

//
// Logo
//
procedure TMainForm.LogoClick(Sender: TObject);
begin
  ShellExecute(Handle,'OPEN','http://gunayato.free.fr',Nil,Nil,SW_SHOW);
end;



//
// Search / Replace grid
//
procedure TMainForm.SRAddBtnClick(Sender: TObject);
begin
  SRGrid.InsertRow(SRGrid.Row+1);
  SRGrid.FixedRows := 1;
end;


procedure TMainForm.SRDelBtnClick(Sender: TObject);
begin
  SRGrid.RemoveRow(SRGrid.Row);
  if SRGrid.RowCount <= 1 then SRGrid.RowCount := 2;
  if SRGrid.RowCount > 1 then SRGrid.FixedRows := 1;
end;



procedure TMainForm.SRExcelExportBtnClick(Sender: TObject);
var
  ExcelFileName: string;
begin
  ExcelFileName := PrjPath + Format('%s_SRGrid', [PrjName]);
  SaveAsExcelFile(SRGrid, PrjName, ExcelFileName);
  ShellExecute(Handle, 'open', PChar(ExcelFileName), nil, nil, SW_SHOW);
end;


procedure TMainForm.SRExcelImportBtnClick(Sender: TObject);
var
  ExcelFileName: string;
begin
  ExcelFileName := PrjPath + Format('%s_SRGrid', [PrjName]);
  if FileExists(ExcelFileName + '.xlsx') then
    ExcelFileName := ExcelFileName + '.xlsx'
  else
    ExcelFileName := ExcelFileName + '.xls';
  ImportExcelFile2StringGrid(SRGrid, ExcelFileName);
  SRGridInit;
end;



procedure TMainForm.SRCopyBtnClick(Sender: TObject);
begin
  SRCopyCol := SRGrid.Col;
  ShowStatus('Column copied.');
end;


procedure TMainForm.SRPasteBtnClick(Sender: TObject);
var
  ColSave: string;
begin
  ColSave := SRGrid.Cells[SRGrid.Col, 0];
  SRGrid.Cols[SRGrid.Col] := SRGrid.Cols[SRCopyCol];
  SRGrid.Cells[SRGrid.Col, 0] := ColSave;
  ShowStatus('Column pasted.');
end;



procedure TMainForm.SRGridDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  with (Sender as TStringGrid) do
  begin
    // Draw the Band
    if (ACol div 2) mod 2 = 0 then begin
      Canvas.Font.Color := clBlack;
      Canvas.Brush.Color := $00E1FFF9;
    end
    else begin
      Canvas.Brush.Color := $00FFEBDF;
      Canvas.Font.Color := clBlue;
    end;
    Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, cells[acol, arow]);
    Canvas.FrameRect(Rect);
  end;
end;



procedure TMainForm.SRGridInit;
begin
  SRGrid.ColCount := 10 * 2;

  if SRGrid.RowCount <= 1 then SRGrid.RowCount := 2;
  if SRGrid.RowCount > 1 then SRGrid.FixedRows := 1;

  SRGrid.Cells[0, 0]  := 'Search 1';
  SRGrid.Cells[1, 0]  := 'Replace 1';
  SRGrid.Cells[2, 0]  := 'Search 2';
  SRGrid.Cells[3, 0]  := 'Replace 2';
  SRGrid.Cells[4, 0]  := 'Search 3';
  SRGrid.Cells[5, 0]  := 'Replace 3';
  SRGrid.Cells[6, 0]  := 'Search 4';
  SRGrid.Cells[7, 0]  := 'Replace 4';
  SRGrid.Cells[8, 0]  := 'Search 5';
  SRGrid.Cells[9, 0]  := 'Replace 5';
  SRGrid.Cells[10, 0] := 'Search 6';
  SRGrid.Cells[11, 0] := 'Replace 6';
  SRGrid.Cells[12, 0] := 'Search 7';
  SRGrid.Cells[13, 0] := 'Replace 7';
  SRGrid.Cells[14, 0] := 'Search 8';
  SRGrid.Cells[15, 0] := 'Replace 8';
  SRGrid.Cells[16, 0] := 'Search 9';
  SRGrid.Cells[17, 0] := 'Replace 9';
  SRGrid.Cells[18, 0] := 'Search 10';
  SRGrid.Cells[19, 0] := 'Replace 10';

  AutoSizeAllCol(SRGrid);

  SRCopyCol := -1;
end;



procedure TMainForm.Button1Click(Sender: TObject);
begin
  AutoSizeAllCol(SRGrid);
end;





//
// Process text
//
procedure TMainForm.ProcessTextByLoopBtnClick(Sender: TObject);
var
  I, J: integer;
  sText, sSearch, sReplace: string;
  bMod: boolean;
begin
  if SRGrid.RowCount < 2 then Exit;

  FinalTextEdit.Text := '';

  // Header
  FinalTextEdit.AddFormatText(HeaderTextEdit.Text);

  for I := 1 to SRGrid.RowCount - 1 do
  begin
    sText := Text2LoopEdit.Text;
    bMod := False;
    for J := 0 to 10 do
    begin
      sSearch := SRGrid.Cells[J*2, I];
      sReplace := SRGrid.Cells[(J*2)+1, I];
      if (sSearch = '') then Continue;
      sText := StringReplace(sText, sSearch, sReplace, [rfReplaceAll]);
      bMod := True;
    end;
    if bMod then
      FinalTextEdit.AddFormatText(sText);
  end;

  // Footer
  FinalTextEdit.AddFormatText(FooterTextEdit.Text);

end;



procedure TMainForm.ProcessTextNormallyBtnClick(Sender: TObject);
var
  I, J: integer;
  sText, sSearch, sReplace: string;
begin
  if SRGrid.RowCount < 2 then Exit;

  FinalTextEdit.Text := '';

  // Header
  FinalTextEdit.AddFormatText(HeaderTextEdit.Text);

  sText := Text2LoopEdit.Text;
  for I := 1 to SRGrid.RowCount - 1 do
  begin
    for J := 0 to 10 do
    begin
      sSearch := SRGrid.Cells[J*2, I];
      sReplace := SRGrid.Cells[(J*2)+1, I];
      if (sSearch = '') then Continue;
      sText := StringReplace(sText, sSearch, sReplace, [rfReplaceAll]);
    end;
  end;
  FinalTextEdit.AddFormatText(sText);

  // Footer
  FinalTextEdit.AddFormatText(FooterTextEdit.Text);

end;




procedure TMainForm.ProjectSheetContextPopup(Sender: TObject; MousePos: TPoint;
  var Handled: Boolean);
begin

end;




//
// Copy final text to clipboard
//
procedure TMainForm.CopyClpBtnClick(Sender: TObject);
begin
  FinalTextEdit.SelectAll;
  FinalTextEdit.CopyToClipboard;
  ShowStatus('Text copied to clipboard.');
end;






//
// Save final text
//
procedure TMainForm.Save2FileBtnClick(Sender: TObject);
begin
  if SaveTextDialog.Execute then
    FinalTextEdit.Lines.SaveToFile(SaveTextDialog.FileName);
end;







//
// Open/Save project
//
procedure TMainForm.SavePrjBtnClick(Sender: TObject);
begin
  if SavePrjDialog.Execute then begin
    try
      PrjName := ChangeFileExt(ExtractFileName(SavePrjDialog.FileName), '');
      PrjPath := ExtractFilePath(SavePrjDialog.FileName);
      SavePrjDialog.InitialDir := PrjPath;
      OpenPrjDialog.InitialDir := PrjPath;

      PrjFileStorage.FileName := PrjPath + PrjName + '.Prj';
      PrjFileStorage.Flush;

      HeaderTextEdit.Lines.SaveToFile(PrjPath + PrjName + '.Header');
      Text2LoopEdit.Lines.SaveToFile(PrjPath + PrjName + '.Text');
      FooterTextEdit.Lines.SaveToFile(PrjPath + PrjName + '.Footer');
      FinalTextEdit.Lines.SaveToFile(PrjPath + PrjName + '.Final');

      SRGrid.SaveToCSV(PrjPath + PrjName + '.SRGrid');

      ShowStatusPrj(PrjName);
      ShowStatus('Project saved.');
    except
      ShowStatus('Project saving error.');
    end;
  end;
end;



procedure TMainForm.LoadPrjBtnClick(Sender: TObject);
var
  Err: boolean;
begin
  Err := False;
  if OpenPrjDialog.Execute then begin
    PrjName := ChangeFileExt(ExtractFileName(OpenPrjDialog.FileName), '');
    PrjPath := ExtractFilePath(OpenPrjDialog.FileName);
    SavePrjDialog.FileName := ExtractFileName(OpenPrjDialog.FileName);
    SavePrjDialog.InitialDir := PrjPath;
    OpenPrjDialog.InitialDir := PrjPath;

    PrjFileStorage.FileName := PrjPath + PrjName + '.Prj';
    try PrjFileStorage.Reload except Err := True; end;

    try HeaderTextEdit.Lines.LoadFromFile(PrjPath + PrjName + '.Header');  except Err := True; end;
    try Text2LoopEdit.Lines.LoadFromFile(PrjPath + PrjName + '.Text');    except Err := True; end;
    try FooterTextEdit.Lines.LoadFromFile(PrjPath + PrjName + '.Footer');  except Err := True; end;
    try FinalTextEdit.Lines.LoadFromFile(PrjPath + PrjName + '.Final');  except Err := True; end;

    try SRGrid.LoadFromCSV(PrjPath + PrjName + '.SRGrid');  except Err := True; end;

    SRGridInit;

    ShowStatusPrj(PrjName);

    if not Err Then
      ShowStatus('Project loaded.')
    else
      ShowStatus('Project error.');
  end;
end;




procedure TMainForm.NewPrjBtnClick(Sender: TObject);
begin
  PrjName := '';
  PrjPath := '';
  PrjFileStorage.FileName := '';
  HeaderTextEdit.Lines.Clear;
  Text2LoopEdit.Lines.Clear;
  FooterTextEdit.Lines.Clear;
  FinalTextEdit.Lines.Clear;
  SRGrid.RowCount := 1;
  ShowStatus('New project done.');
end;




////////////////////////////////////////////



//
// Auto size grid's column
//
procedure AutoSizeCol(Grid: TStringGrid; Column: integer);
var
  i, W, WMax: integer;
begin
  WMax := 0;
  for i := 0 to (Grid.RowCount - 1) do begin

    if i = 0 then
      Grid.Canvas.Font.Style := [fsBold]
    else
      Grid.Canvas.Font.Style := [];


    W := Grid.Canvas.TextWidth(Grid.Cells[Column, i]+'X');
    if W > WMax then
      WMax := W;
  end;
  Grid.ColWidths[Column] := WMax;
end;

procedure AutoSizeAllCol(Grid: TStringGrid);
var
  i: integer;
begin
  for i := 0 to Grid.ColCount - 1 do
    AutoSizeCol(Grid, i);
end;







end.
