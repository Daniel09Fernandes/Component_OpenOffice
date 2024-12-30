unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, uniGUITypes, uniGUIAbstractClasses,
  uniGUIClasses, uniGUIRegClasses, uniGUIForm, uOpenOffice, uOpenOffice.Calc,
  uOpenOffice.Helpers, uOpenOffice.Collors,
  uniGUIBaseClasses, uniButton, uniFileUpload, uOpenOffice.Writer;

type
  TMainForm = class(TUniForm)
    UniButton1: TUniButton;
    up: TUniFileUpload;
    UniButton2: TUniButton;
    UniButton3: TUniButton;
    OpenOffice_writer1: TOpenOffice_writer;
    UniButton4: TUniButton;
    UniButton5: TUniButton;
    UniButton6: TUniButton;
    procedure UniButton1Click(Sender: TObject);
    procedure UniButton2Click(Sender: TObject);
    procedure upCompleted(Sender: TObject; AStream: TFileStream);
    procedure UniButton3Click(Sender: TObject);
    procedure UniButton4Click(Sender: TObject);
    procedure UniButton5Click(Sender: TObject);
  private
    OpenOffice_calc1 : TOpenOffice_calc;
  public
    { Public declarations }
  end;

function MainForm: TMainForm;

implementation

{$R *.dfm}

uses
  uniGUIVars, MainModule, uniGUIApplication, ServerModule;

function MainForm: TMainForm;
begin
  Result := TMainForm(UniMainModule.GetFormInstance(TMainForm));
end;

procedure TMainForm.UniButton1Click(Sender: TObject);
var SettingsChart: TSettingsChart;
begin
  UniMainModule.OpenOffice_calc1.DocVisible := true;
  UniMainModule.OpenOffice_calc1.startSheet;

  UniMainModule.OpenOffice_calc1.SetValue(1, 'A', 'STATUS')
    .SetBorder([bAll], opBrown)
    .changeJustify(fthRIGHT, ftvTOP)
    .setBold(true)
    .changeFont('Arial', 12)
    .SetUnderline(true)
    .setColor(opWhite, opMagenta);

  UniMainModule.OpenOffice_calc1.SetValue(1, 'B', 'VALOR').changeJustify(fthRIGHT, ftvTOP)
    .SetBorder([bAll], opBrown)
    .setBold(true)
    .changeFont('Arial', 12)
    .SetUnderline(true)
    .setColor(opWhite, opMagenta);

  UniMainModule.OpenOffice_calc1.SetValue(2, 'B', 109, ftNumeric).SetBorder([bAll], opBrown);
  UniMainModule.OpenOffice_calc1.SetValue(2, 'A', 'AGUA').SetBorder([bAll], opBrown);

  UniMainModule.OpenOffice_calc1.SetValue(3, 'B', 105.55, ftNumeric).SetBorder([bAll], opBrown);
  UniMainModule.OpenOffice_calc1.SetValue(3, 'A', 'LUZ').SetBorder([bAll], opBrown);

  UniMainModule.OpenOffice_calc1.SetValue(4, 'B', 1005.22, ftNumeric);
  UniMainModule.OpenOffice_calc1.SetValue(4, 'A', 'ALUGUEL');

  UniMainModule.OpenOffice_calc1.SetValue(6, 'A', 'Total de linhas');
  UniMainModule.OpenOffice_calc1.SetValue(6, 'B', UniMainModule.OpenOffice_calc1.CountRow, ftNumeric);

  UniMainModule.OpenOffice_calc1.SetValue(7, 'A', 'Total de Colunas');
  UniMainModule.OpenOffice_calc1.SetValue(7, 'B', UniMainModule.OpenOffice_calc1.CountCell, ftNumeric);

  UniMainModule.OpenOffice_calc1.addNewSheet('A Receber', 1);

  UniMainModule.OpenOffice_calc1.SetValue(1, 'A', 'VALOR')
    .SetBorder([bAll], opBrown)
    .changeJustify(fthRIGHT, ftvTOP)
    .setBold(true);

  UniMainModule.OpenOffice_calc1.SetValue(1, 'B', 'DESC')
    .SetBorder([bAll], opBrown)
    .changeJustify(fthRIGHT, ftvTOP)
    .setBold(true)
    .changeFont('Arial', 12)
    .SetUnderline(true)
    .setColor(opWhite, opCiano);

  UniMainModule.OpenOffice_calc1.SetValue(1, 'C', 'SOMA')
    .SetBorder([bAll], opBrown)
    .changeJustify(fthRIGHT, ftvTOP)
    .setBold(true)
    .changeFont('Arial', 12)
    .SetUnderline(true)
    .setColor(opWhite, opSoftRed);

  UniMainModule.OpenOffice_calc1.SetValue(1, 'H', 'SOMA')
    .SetBorder([bAll], opBrown)
    .changeJustify(fthRIGHT, ftvTOP)
    .setBold(true)
    .changeFont('Arial', 12)
    .SetUnderline(true)
    .setColor(opWhite, opSoftRed);

  UniMainModule.OpenOffice_calc1.SetValue(2, 'A', 200, ftNumeric);
  UniMainModule.OpenOffice_calc1.SetValue(2, 'B', 'Emprestimo');
  UniMainModule.OpenOffice_calc1.SetValue(2, 'C', 0, ftNumeric);

  UniMainModule.OpenOffice_calc1.SetValue(3, 'A', 369.55, ftNumeric);
  UniMainModule.OpenOffice_calc1.SetValue(3, 'B', 'Dividendos');
  UniMainModule.OpenOffice_calc1.SetValue(3, 'C', 0, ftNumeric);

  UniMainModule.OpenOffice_calc1.SetValue(4, 'A', 1585.22, ftNumeric);
  UniMainModule.OpenOffice_calc1.SetValue(4, 'B', 'ALUGUEL');
  UniMainModule.OpenOffice_calc1.SetValue(4, 'C', 0, ftNumeric);

  UniMainModule.OpenOffice_calc1.SetValue(8, 'A', 1585.22, ftNumeric);
  UniMainModule.OpenOffice_calc1.SetValue(8, 'B', 'Renda extra');
  UniMainModule.OpenOffice_calc1.SetValue(8, 'C', 0, ftNumeric);

  UniMainModule.OpenOffice_calc1.SetValue(15, 'A', 1585.22, ftNumeric);
  UniMainModule.OpenOffice_calc1.SetValue(15, 'B', 'ALUGUEL 2');
  UniMainModule.OpenOffice_calc1.SetValue(15, 'C', 0, ftNumeric);

  UniMainModule.OpenOffice_calc1.SetValue(17, 'A', 'Total de linhas');
  UniMainModule.OpenOffice_calc1.SetValue(17, 'B', UniMainModule.OpenOffice_calc1.CountRow, ftNumeric);

  UniMainModule.OpenOffice_calc1.SetValue(19, 'A', 'Total de Colunas');
  UniMainModule.OpenOffice_calc1.SetValue(19, 'B', UniMainModule.OpenOffice_calc1.CountCell, ftNumeric);
  UniMainModule.OpenOffice_calc1.setFormula(20, 'A', '=A2+A3+A4+A15')
    .setBold(true);

  UniMainModule.OpenOffice_calc1.positionSheetByName('Planilha1');

  //Configure the chart settings
 
  SettingsChart.Height := 11000;
  SettingsChart.Width := 22000;
  SettingsChart.Position_X := 1500;
  SettingsChart.Position_Y := 5000;
  SettingsChart.StartRow := 0;
  SettingsChart.EndRow := 3;
  SettingsChart.PositionSheet := 0; //first tab
  SettingsChart.StartColumn := 'A';
  SettingsChart.EndColumn := 'B';
  SettingsChart.ChartName := 'TestChart';
  SettingsChart.typeChart := ctDefault;

  UniMainModule.OpenOffice_calc1.addChart(SettingsChart);

  SettingsChart.typeChart := ctVertical;
  UniMainModule.OpenOffice_calc1.addChart(SettingsChart);

  SettingsChart.typeChart := ctPie;
  UniMainModule.OpenOffice_calc1.addChart(SettingsChart);

  SettingsChart.typeChart := ctLine;
  UniMainModule.OpenOffice_calc1.addChart(SettingsChart);

  sleep(3000);
  UniMainModule.OpenOffice_calc1.saveFile(UniServerModule.TempFolderPath+'sheet.xls');
  UniMainModule.OpenOffice_calc1.CloseFile;

  UniSession.SendFile(UniServerModule.TempFolderPath+'sheet.xls' );
end;

procedure TMainForm.UniButton2Click(Sender: TObject);
begin
  up.Execute;
end;

procedure TMainForm.UniButton3Click(Sender: TObject);
begin
  UniMainModule.OpenOffice_calc1.CloseFile;
end;

procedure TMainForm.UniButton4Click(Sender: TObject);
begin
  OpenOffice_writer1.DocVisible := true;
  OpenOffice_writer1.startDoc;
  OpenOffice_writer1.setBold(true)
    .setFontHeight(16)
    .setValue('Título: Apresentando o componente Libre Office writer via Delphi <3' + #13#13);
  OpenOffice_writer1.gotoEndOfSentence;

  OpenOffice_writer1.setBold(false)
    .setFontHeight(12)
    .setValue('Neste exemplo estou mostrando a criação de documentos via código, de um jeito simples, rápido e fácil.' + #13);
  OpenOffice_writer1.gotoEndOfSentence;

  OpenOffice_writer1.setBold(false)
    .setFontHeight(12)
    .setValue('Espero que seja útil e que estejam gostando, lembrando o componente é open source e totalmente free!' + #13#13);
  OpenOffice_writer1.gotoEndOfSentence;

  OpenOffice_writer1.setBold(false)
    .setFontHeight(12)
    .SetBold(true)
    .setValue('Obrigado a todos pela presença!!!' + #13);
  OpenOffice_writer1.gotoEndOfSentence;

  sleep(3000);
  OpenOffice_writer1.saveFile(UniServerModule.TempFolderPath+'doc.docx');
  OpenOffice_writer1.CloseFile;

  UniSession.SendFile(UniServerModule.TempFolderPath+'doc.docx' );
end;

procedure TMainForm.UniButton5Click(Sender: TObject);
begin
  OpenOffice_writer1.CloseFile;
end;

procedure TMainForm.upCompleted(Sender: TObject; AStream: TFileStream);
begin
  if not String(up.FileName).Trim.IsEmpty  then
  begin
    if assigned(AStream) then
     AStream.free;

     if Pos('xls', up.FileName) > 0 then
     begin
       UniMainModule.OpenOffice_calc1.DocVisible  := true;
       UniMainModule.OpenOffice_calc1.URlFile     := up.Files[0].CacheFile;
       UniMainModule.OpenOffice_calc1.SheetName   := 'Planilha1';
       UniMainModule.OpenOffice_calc1.startSheet;
     end else
     begin
        OpenOffice_writer1.DocVisible  := true;
       OpenOffice_writer1.URlFile     := up.Files[0].CacheFile;
       OpenOffice_writer1.startDoc;
     end;

  end;
end;

initialization

RegisterAppFormClass(TMainForm);

end.
