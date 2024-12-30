{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOfficeHelper.pas }
{ Developer: Daniel Fernandes Rodrigures }
{ Email    : danielfernandesroddrigues@gmail.com }
{ this unit is a part of the Open Source. }
{ licensed under a MIT - see LICENSE.md}

{ ******************************************************* }

unit uOpenOffice.Helpers;

interface

uses vcl.stdCtrls, System.SysUtils, math, System.Variants,
 uOpenOffice.Calc,
 uOpenOffice.Writer,
 uOpenOffice.Collors;

type
  TBorder = (bAll, bLeft, bRight, bBottom, bTop);

  TBoderSheet = set of TBorder;

  { STANDARD : é o alinhamento padrão tanto para números como para textos, sendo a esqueda para as strings e a direita para os números;
    LEFT : o conteúdo é alinhado no lado esquerdo da célula;
    CENTER : o conteúdo é alinhado no centro da célula;
    RIGHT : o conteúdo é alinhado no lado direito da célula;
    BLOCK : o conteúdo é alinhando em relação ao comprimento da célula;
    REPEAT : o conteúdo é repetido dentro da célula para preenchê-la. }
  THoriJustify = (fthSTANDARD, fthLEFT, fthCENTER, fthRIGHT, fthBLOCK,
    fthREPEAT);
  { STANDARD : é o valor usado como padrão;
    TOP : o conteúdo da célula é alinhado pelo topo;
    CENTER : o conteúdo da célula é alinhado pelo centro;
    BOTTOM : o conteúdo da célula é alinhado pela base. }
  TVertJustify = (ftvSTANDARD, ftvTOP, ftvCENTER, ftvBOTTOM);

  TTypeChart = (ctDefault, ctVertical, ctPie, ctLine);

  THelperHoriJustify = record Helper for THoriJustify
  public
    function ToInteger: Integer;
  end;

  THelperVertJustify = record Helper for TVertJustify
  public
    function ToInteger: Integer;
  end;

  TSettingsChart = record
    Height,
    Width,
    Position_X,
    Position_Y,
    StartRow,
    PositionSheet,
    EndRow: Integer;
    StartColumn,
    EndColumn,
    ChartName: string;
    TypeChart: TTypeChart;
  end;

  THelperOpenOffice_Writer = class Helper for TOpenOffice_Writer
    function SetUnderline(AUnderline: Boolean): TOpenOffice_Writer;
    function SetBold(ABold: Boolean): TOpenOffice_Writer;
    function SetFontHeight(AFontHeight: Integer): TOpenOffice_Writer;
    function SetColorText(AColor: TOpenColor) : TOpenOffice_Writer;
    function SetFontName(AFont : string): TOpenOffice_Writer;
  end;

  THelperOpenOffice_Calc = class Helper for TOpenOffice_Calc
    procedure AddChart(ASettingsChart: TSettingsChart);
    function SetBorder(ABorderPosition: TBoderSheet; AOpColor: TOpenColor; ARemoveBorder: Boolean = False) : TOpenOffice_Calc;
    function ChangeFont(AFontName: string; AHeight: Integer): TOpenOffice_Calc;
    function ChangeJustify(ATypeHori: THoriJustify; ATypeVert: TVertJustify) : TOpenOffice_Calc;
    function SetColor(AFontColor, ABackgroud: TOpenColor): TOpenOffice_Calc;
    function SetCellWidth(const AWidth: Integer): TOpenOffice_Calc;
    function SetBold(ABold: Boolean): TOpenOffice_Calc;
    function SetUnderline(AUnderline: Boolean): TOpenOffice_Calc;
    function CountRow: Integer;
    function CountCell: Integer;
    function SheetToBase64(APathFile:string):string;
  end;

implementation

uses
  System.Win.ComObj, System.Classes, Soap.EncdDecd;

procedure THelperOpenOffice_Calc.AddChart(ASettingsChart: TSettingsChart);
var
  lChart, lRect, lSheet : OleVariant;
  lRangeAddress: Variant;
  lCountChart: Integer;
begin
  lCountChart := 1;

  if ASettingsChart.ChartName.Trim.IsEmpty then
    ASettingsChart.ChartName := 'MyChart_' + (ASettingsChart.StartColumn + ASettingsChart.StartRow.ToString) + '_' +
      (ASettingsChart.EndColumn + ASettingsChart.EndRow.ToString);

  lSheet := ObjDocument.Sheets.GetByIndex(ASettingsChart.PositionSheet);
  // getByName(aCollName);
  ObjCharts := lSheet.Charts;

  while ObjCharts.hasByName(ASettingsChart.ChartName) do
  begin
    ASettingsChart.ChartName := copy(ASettingsChart.ChartName,0, ifthen( (pos('_',ASettingsChart.ChartName) > 0),
                                               pos('_',ASettingsChart.ChartName), ASettingsChart.ChartName.Length)
                      ) + '_' + lCountChart.ToString;
    Inc(lCountChart);
    ASettingsChart.Position_Y := (ASettingsChart.Position_Y + ASettingsChart.Height) + 1000;
  end;

  lRect := ObjServiceManager.Bridge_GetStruct('com.sun.star.awt.Rectangle');
  lRangeAddress := lSheet.Bridge_GetStruct('com.sun.star.table.CellRangeAddress');

  lRect.Width := ASettingsChart.Width;
  lRect.Height := ASettingsChart.Height;
  lRect.X := ASettingsChart.Position_X;
  lRect.Y := ASettingsChart.Position_Y;

  lRangeAddress.Sheet := ASettingsChart.PositionSheet;
  lRangeAddress.StartColumn := Fields.getIndex(ASettingsChart.StartColumn);
  lRangeAddress.StartRow := ASettingsChart.StartRow;
  lRangeAddress.EndColumn := Fields.getIndex(ASettingsChart.EndColumn);
  lRangeAddress.EndRow := ASettingsChart.EndRow;

  ObjCharts.addNewByName(ASettingsChart.ChartName, lRect, VarArrayOf(lRangeAddress), True, True);

  if ASettingsChart.TypeChart <> ctDefault then
  begin
    lChart := ObjCharts.getByName(ASettingsChart.ChartName).embeddedObject;
    lChart.Title.String := ASettingsChart.ChartName;
    case ASettingsChart.TypeChart of
      ctVertical:
        lChart.Diagram.Vertical := True;
      ctPie:
        begin
          lChart.Diagram := lChart.createInstance
            ('com.sun.star.chart.PieDiagram');
          lChart.HasMainTitle := True;
        end;
      ctLine:
        begin
          lChart.Diagram := lChart.createInstance
            ('com.sun.star.chart.LineDiagram');
        end;
    end;
  end;

end;

function THelperOpenOffice_Calc.ChangeFont(AFontName: string; AHeight: Integer)
  : TOpenOffice_Calc;
begin
  // Cell := Table.getCellRangeByName(aCollName+aCellNumber.ToString);
  if not AFontName.Trim.IsEmpty then
    Cell.CharFontName := AFontName;

  Cell.CharHeight := IntToStr(AHeight);
  Result := Self;
end;

function THelperOpenOffice_Calc.ChangeJustify(ATypeHori: THoriJustify;
  ATypeVert: TVertJustify): TOpenOffice_Calc;
begin
  Cell.HoriJustify := ATypeHori.ToInteger;
  Cell.VertJustify := ATypeVert.ToInteger;
  Result := Self;
end;

function THelperOpenOffice_Calc.CountRow: Integer;
var
  lRow, lCountRow: Integer;
  lCountBlank: Integer;
  lBreak, lAllBlank: Boolean;
  lIdx: Integer;
begin
  lBreak := False;
  lRow := 1;
  lCountRow := 0;
  lCountBlank := 0;

  while not lBreak do
  begin
    for lIdx := 0 to 21 do
    begin
      if GetValue(lRow, Fields.getField(lIdx)).Value.Trim.IsEmpty then
      begin
        lAllBlank := True;
      end
      else
      begin
        if lCountBlank > 0 then // An empty column behind a valued column
          lCountRow := lCountRow + lCountBlank;

        lAllBlank := False;
        lCountBlank := 0;

        Inc(lCountRow);
        break;
      end;
    end;
    Inc(lRow);

    if lCountBlank = 50 then
      lBreak := True;

    if lAllBlank then
      Inc(lCountBlank);

  end;
  Result := lCountRow;
end;

function THelperOpenOffice_Writer.SetBold(ABold: Boolean): TOpenOffice_Writer;
var lCtrlBold: Boolean;
begin
  lCtrlBold := False;
  PropsText[0].Name := 'bold';
  PropsText[0].Value := ABold;
  //Funcionando, porém rever
  if BoldActive and (not ABold) then
  begin
    BoldActive   := False;
    lCtrlBold     := True;
    ABold        := True;
  end;

  if ( not BoldActive) and (ABold) then
  begin
    ObjDispatcher.ExecuteDispatch(ObjWriter, '.uno:Bold', '', 0,  VarArrayOf(PropsText));

    if not lCtrlBold then
      BoldActive := True;
  end;

  Result := Self;
end;

function THelperOpenOffice_Writer.SetFontName(AFont: string): TOpenOffice_Writer;
begin
  if not ValueText.Trim.IsEmpty then
  begin
    Cursor.CharFontName := AFont;
    SetValue(ValueText);
  end;
  Result := Self;
end;

function THelperOpenOffice_Writer.SetColorText(AColor: TOpenColor): TOpenOffice_Writer;
begin
  if not ValueText.Trim.IsEmpty then
  begin
     Cursor.SetPropertyValue('CharColor', AColor);
     SetValue(ValueText);
  end;

  Result := Self;
end;

function THelperOpenOffice_Writer.SetFontHeight(AFontHeight: Integer) : TOpenOffice_Writer;
begin
  PropsText[1].Name := 'FontHeight.Height';
  PropsText[1].Value := AFontHeight;
  ObjDispatcher.ExecuteDispatch(ObjWriter, '.uno:FontHeight', '', 0, VarArrayOf(PropsText));

  Result := Self;
end;

function THelperOpenOffice_Writer.SetUnderline(AUnderline: Boolean): TOpenOffice_Writer;
begin
  if not ValueText.Trim.IsEmpty then
  begin
    Cursor.CharUnderline := ifthen(AUnderline,1,0);
    SetValue(ValueText);
  end;

  Result := Self;
end;

function THelperOpenOffice_Calc.CountCell: Integer;
var
  lCell, lCountCell, lCountBlank: Integer;
  lIdx: Integer;
  lAllBlank: Boolean;
begin
  lCell := 1;
  lCountCell := 0;
  lCountBlank := 0;

  for lIdx := 0 to 21 do
  begin
    for lCell := 1 to 10 do
    begin
      if not GetValue(lCell, Fields.getField(lIdx)).Value.Trim.IsEmpty then
      begin

        if lCountBlank > 0 then
          lCountCell := lCountCell + lCountBlank;

        lAllBlank := False;

        Inc(lCountCell);
        break;
      end
      else
        lAllBlank := True;
    end;

    if lCountBlank = 10 then
    begin
      lCountBlank := 0;
      break;
    end;

    if lAllBlank then
      Inc(lCountBlank);
  end;

  Result := lCountCell;
end;

function THelperOpenOffice_Calc.SetBorder(ABorderPosition: TBoderSheet; AOpColor: TOpenColor; ARemoveBorder: Boolean): TOpenOffice_Calc;
var
  lSettings: Variant;
begin
 CoreReflection.forName('com.sun.star.table.BorderLine2').createObject(lSettings);

 if not ARemoveBorder then
  begin
    lSettings.Color := AOpColor;
    lSettings.InnerLineWidth := 20;
    lSettings.LineDistance := 60;
    lSettings.LineWidth := 2;
    lSettings.OuterLineWidth := 20;
  end else
  begin
    lSettings.Color := 0;
    lSettings.InnerLineWidth := 0;
    lSettings.LineDistance := 0;
    lSettings.LineWidth := 0;
    lSettings.OuterLineWidth := 0;
  end;

  if bAll in ABorderPosition then
  begin
    Cell.TopBorder := lSettings;
    Cell.LeftBorder := lSettings;
    Cell.RightBorder := lSettings;
    Cell.BottomBorder := lSettings;
  end;

  if bTop in ABorderPosition then
    Cell.TopBorder := lSettings;

  if bLeft in ABorderPosition then
    Cell.LeftBorder := lSettings;

  if bRight in ABorderPosition then
    Cell.RightBorder := lSettings;

  if bBottom in ABorderPosition then
    Cell.BottomBorder := lSettings;

  Result := Self;
end;

function THelperOpenOffice_Calc.SetCellWidth(const AWidth: Integer): TOpenOffice_Calc;
begin
   Cell.GetColumns.GetByIndex(0).Width := AWidth;
end;

function THelperOpenOffice_Calc.SetColor(AFontColor, ABackgroud: TOpenColor)
  : TOpenOffice_Calc;
begin
  Cell.CharColor := AFontColor;
  Cell.CellBackColor := ABackgroud;
  Result := Self;
end;

function THelperOpenOffice_Calc.SetBold(ABold: Boolean): TOpenOffice_Calc;
begin
  Cell.CharWeight := ifthen(ABold, 150, 0);
  Result := Self;
end;

function THelperOpenOffice_Calc.SetUnderline(AUnderline: Boolean): TOpenOffice_Calc;
begin
  Cell.CharUnderline := ifthen(AUnderline, 1, 0);
  Result := Self;
end;

function THelperOpenOffice_Calc.SheetToBase64(APathFile:string): string;
var
  lStream: TMemoryStream;
begin
  lStream := TMemoryStream.Create;
  try
    lStream.LoadFromFile(APathFile);
    Result := EncodeBase64(lStream.Memory, lStream.Size);
  finally
    lStream.Free;
  end;
end;

{ THelperOpenOffice_Calc }

function THelperHoriJustify.ToInteger: Integer;
begin
  case Self of
    fthSTANDARD:
      Result := 0;
    fthLEFT:
      Result := 1;
    fthCENTER:
      Result := 2;
    fthRIGHT:
      Result := 3;
    fthBLOCK:
      Result := 4;
    fthREPEAT:
      Result := 5;
    else
      raise Exception.Create('Unknown Justification Value');
  end;
end;

{ THelperVertJustify }

function THelperVertJustify.ToInteger: Integer;
begin
  case Self of
    ftvSTANDARD:
      Result := 0;
    ftvTOP:
      Result := 1;
    ftvCENTER:
      Result := 2;
    ftvBOTTOM:
      Result := 3;
    else
      raise Exception.Create('Unknown Justification Value');
  end;
end;

end.
