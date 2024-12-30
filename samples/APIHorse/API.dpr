program API;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  System.SysUtils,
  Horse,
  uOpenOffice.Calc,
  uOpenOffice.Helpers,
  uOpenOffice.Collors;

procedure GetSheet(Req: THorseRequest; Res: THorseResponse; Next: TProc);
begin
  var
  OpenOffice_calc1 := TOpenOffice_calc.Create(nil);
  try
    OpenOffice_calc1.DocVisible := false;
    OpenOffice_calc1.StartSheet;

    OpenOffice_calc1.SetValue(1, 'A', 'STATUS').SetBorder([bAll], opBrown)
      .changeJustify(fthRIGHT, ftvTOP).setBold(true).changeFont('Arial', 12)
      .SetUnderline(true).setColor(opWhite, opMagenta);

    OpenOffice_calc1.SetValue(1, 'B', 'VALOR').changeJustify(fthRIGHT, ftvTOP)
      .SetBorder([bAll], opBrown).setBold(true).changeFont('Arial', 12)
      .SetUnderline(true).setColor(opWhite, opMagenta);

    OpenOffice_calc1.SetValue(2, 'B', 109, ftNumeric)
      .SetBorder([bAll], opBrown);
    OpenOffice_calc1.SetValue(2, 'A', 'AGUA').SetBorder([bAll], opBrown);

    OpenOffice_calc1.SetValue(3, 'B', 105.55, ftNumeric)
      .SetBorder([bAll], opBrown);
    OpenOffice_calc1.SetValue(3, 'A', 'LUZ').SetBorder([bAll], opBrown);

    OpenOffice_calc1.SetValue(4, 'B', 1005.22, ftNumeric);
    OpenOffice_calc1.SetValue(4, 'A', 'ALUGUEL');

    OpenOffice_calc1.SetValue(6, 'A', 'Total de linhas');
    OpenOffice_calc1.SetValue(6, 'B', OpenOffice_calc1.CountRow, ftNumeric);

    OpenOffice_calc1.SetValue(7, 'A', 'Total de Colunas');
    OpenOffice_calc1.SetValue(7, 'B', OpenOffice_calc1.CountCell, ftNumeric);

    OpenOffice_calc1.AddNewSheet('A Receber', 1);

    OpenOffice_calc1.SetValue(1, 'A', 'VALOR').SetBorder([bAll], opBrown)
      .changeJustify(fthRIGHT, ftvTOP).setBold(true);

    OpenOffice_calc1.SetValue(1, 'B', 'DESC').SetBorder([bAll], opBrown)
      .changeJustify(fthRIGHT, ftvTOP).setBold(true).changeFont('Arial', 12)
      .SetUnderline(true).setColor(opWhite, opCiano);

    OpenOffice_calc1.SetValue(1, 'C', 'SOMA').SetBorder([bAll], opBrown)
      .changeJustify(fthRIGHT, ftvTOP).setBold(true).changeFont('Arial', 12)
      .SetUnderline(true).setColor(opWhite, opSoftRed);

    OpenOffice_calc1.SetValue(1, 'H', 'SOMA').SetBorder([bAll], opBrown)
      .changeJustify(fthRIGHT, ftvTOP).setBold(true).changeFont('Arial', 12)
      .SetUnderline(true).setColor(opWhite, opSoftRed);

    OpenOffice_calc1.SetValue(2, 'A', 200, ftNumeric);
    OpenOffice_calc1.SetValue(2, 'B', 'Emprestimo');
    OpenOffice_calc1.SetValue(2, 'C', 0, ftNumeric);

    OpenOffice_calc1.SetValue(3, 'A', 369.55, ftNumeric);
    OpenOffice_calc1.SetValue(3, 'B', 'Dividendos');
    OpenOffice_calc1.SetValue(3, 'C', 0, ftNumeric);

    OpenOffice_calc1.SetValue(4, 'A', 1585.22, ftNumeric);
    OpenOffice_calc1.SetValue(4, 'B', 'ALUGUEL');
    OpenOffice_calc1.SetValue(4, 'C', 0, ftNumeric);

    OpenOffice_calc1.SetValue(8, 'A', 1585.22, ftNumeric);
    OpenOffice_calc1.SetValue(8, 'B', 'Renda extra');
    OpenOffice_calc1.SetValue(8, 'C', 0, ftNumeric);

    OpenOffice_calc1.SetValue(15, 'A', 1585.22, ftNumeric);
    OpenOffice_calc1.SetValue(15, 'B', 'ALUGUEL 2');
    OpenOffice_calc1.SetValue(15, 'C', 0, ftNumeric);

    OpenOffice_calc1.SetValue(17, 'A', 'Total de linhas');
    OpenOffice_calc1.SetValue(17, 'B', OpenOffice_calc1.CountRow, ftNumeric);

    OpenOffice_calc1.SetValue(19, 'A', 'Total de Colunas');
    OpenOffice_calc1.SetValue(19, 'B', OpenOffice_calc1.CountCell, ftNumeric);
    OpenOffice_calc1.SetFormula(20, 'A', '=A2+A3+A4+A15').setBold(true);

    OpenOffice_calc1.PositionSheetByName('Planilha1');

    // Configure the chart settings
    var
      SettingsChart: TSettingsChart;

    SettingsChart.Height := 11000;
    SettingsChart.Width := 22000;
    SettingsChart.Position_X := 1500;
    SettingsChart.Position_Y := 5000;
    SettingsChart.StartRow := 0;
    SettingsChart.EndRow := 3;
    SettingsChart.PositionSheet := 0; // first tab
    SettingsChart.StartColumn := 'A';
    SettingsChart.EndColumn := 'B';
    SettingsChart.ChartName := 'TestChart';
    SettingsChart.typeChart := ctDefault;

    OpenOffice_calc1.AddChart(SettingsChart);

    SettingsChart.typeChart := ctVertical;
    OpenOffice_calc1.AddChart(SettingsChart);

    SettingsChart.typeChart := ctPie;
    OpenOffice_calc1.AddChart(SettingsChart);

    SettingsChart.typeChart := ctLine;
    OpenOffice_calc1.AddChart(SettingsChart);

    OpenOffice_calc1.SaveFile(GetHomePath + '\sheet.xls');
    OpenOffice_calc1.CloseFile;

    Res.Send(OpenOffice_calc1.SheetToBase64(GetHomePath + '\sheet.xls'));
  finally
    OpenOffice_calc1.Free;
  end;

end;

begin
  try
    THorse.Get('/GetSheet', GetSheet);
    THorse.Listen(9000);
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;

end.
