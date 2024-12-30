unit Sample.View.Form.SampleNoPackage;

interface

uses
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Variants,
  System.Classes,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  math,
  Vcl.Dialogs,
  Vcl.StdCtrls,
  Vcl.ExtCtrls,
  Vcl.Buttons,
  Vcl.Mask,
  Vcl.ComCtrls,
  Data.DB,
  Datasnap.DBClient,
  Vcl.Grids,
  Vcl.DBGrids,
  Vcl.Menus,
  Datasnap.Provider,

  uOpenOffice.Calc,
  UOpenOffice.Writer,
  uOpenOffice.Collors,
  uOpenOffice.Helpers,
  uOpenOffice.SetPrinter,
  uOpenOffice;
type
  TFormSampleNoPackage = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    edtSalvar: TLabeledEdit;
    Button4: TButton;
    edtArq: TLabeledEdit;
    Button5: TButton;
    edtAba: TLabeledEdit;
    edtLinha: TLabeledEdit;
    edtColuna: TLabeledEdit;
    edtValor: TLabeledEdit;
    btAddCellValue: TButton;
    btGetCellValue: TButton;
    chNumeric: TCheckBox;
    cbQuebraLinha: TCheckBox;
    Button8: TButton;
    Button9: TButton;
    edtPos: TLabeledEdit;
    GroupBox1: TGroupBox;
    RBhCenter: TRadioButton;
    RBhLeft: TRadioButton;
    RBhRight: TRadioButton;
    GroupBox2: TGroupBox;
    RBvTop: TRadioButton;
    RBvBottom: TRadioButton;
    RBvCenter: TRadioButton;
    Label1: TLabel;
    CBBold: TCheckBox;
    CBUnderline: TCheckBox;
    CBCorFont: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    CBCorFundo: TComboBox;
    Label4: TLabel;
    cbFontes: TComboBox;
    Graficos: TGroupBox;
    RBDefault: TRadioButton;
    RBVertical: TRadioButton;
    RBPie: TRadioButton;
    RBLine: TRadioButton;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    edtLde: TLabeledEdit;
    edtLAte: TLabeledEdit;
    edtCde: TLabeledEdit;
    edtCAte: TLabeledEdit;
    edtNomeGrafico: TLabeledEdit;
    Button10: TButton;
    Bevel4: TBevel;
    btGetSheetColumnsCount: TBitBtn;
    btGetSheetRowsCount: TBitBtn;
    PageControl1: TPageControl;
    PageControl2: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    BitBtn5: TBitBtn;
    BitBtn6: TBitBtn;
    mmo: TMemo;
    BitBtn7: TBitBtn;
    edtFontHg: TEdit;
    CB_Bold: TCheckBox;
    Label5: TLabel;
    Label6: TLabel;
    CBErase: TCheckBox;
    BitBtn8: TBitBtn;
    cbUnderline_wrt: TCheckBox;
    cbColorWriter: TComboBox;
    BitBtn10: TBitBtn;
    edtTamanhoFonte: TEdit;
    cdsSampleValues: TClientDataSet;
    DBGrid1: TDBGrid;
    dsSampleValues: TDataSource;
    cdsSampleValuesID: TIntegerField;
    cdsSampleValuesNome: TStringField;
    cdsSampleValuesIdade: TIntegerField;
    PopupMenu1: TPopupMenu;
    Exportarplanilha1: TMenuItem;
    edtArqWriter: TLabeledEdit;
    Button11: TButton;
    Button12: TButton;
    lbl1: TLabel;
    lbl2: TLabel;
    CheckBox1: TCheckBox;
    edtWidth: TLabeledEdit;
    Label7: TLabel;
    Button6: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure btAddCellValueClick(Sender: TObject);
    procedure btGetCellValueClick(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure btGetSheetColumnsCountClick(Sender: TObject);
    procedure btGetSheetRowsCountClick(Sender: TObject);
    procedure OpenOffice1BeforeStartSheet(sender: TObject);
    procedure OpenOffice1BeforePrint(sender: TObject; var SetPrinter: TSetPrinter);
    procedure OpenOffice1BeforeGetValue(sender: TObject);
    procedure OpenOffice1BeforeSetValue(sender: TObject);
    procedure OpenOffice1BeforeCloseSheet(sender: TObject);
    procedure OpenOffice1AfterStartSheet(sender: TObject);
    procedure OpenOffice1AfterSetValue(sender: TObject);
    procedure OpenOffice1AfterGetValue(sender: TObject);
    procedure OpenOffice1AfterCloseSheet(sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure BitBtn8Click(Sender: TObject);
    procedure TabSheet2Show(Sender: TObject);
    procedure TabSheet1Show(Sender: TObject);
    procedure BitBtn10Click(Sender: TObject);
    procedure Exportarplanilha1Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
  private
    FFontTop, FFontLeft: integer;
    FSettingsChart: TSettingsChart;

    FOfficeWriter: TOpenOffice_writer;
    FOfficeCalc: TOpenOffice_calc;

    procedure CapturarNomesDeFontes;
    procedure CreateDemoSheet;
    procedure LoadSampleDataSetValues;
    procedure LoadAuxiliarItems;

    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormSampleNoPackage: TFormSampleNoPackage;

implementation

{$R *.dfm}

procedure TFormSampleNoPackage.BitBtn10Click(Sender: TObject);
begin
  FOfficeWriter.gotoEndOfSentence;
end;

procedure TFormSampleNoPackage.btGetSheetColumnsCountClick(Sender: TObject);
begin
  ShowMessage(FOfficeCalc.CountCell.ToString);
end;

procedure TFormSampleNoPackage.btGetSheetRowsCountClick(Sender: TObject);
begin
  ShowMessage(FOfficeCalc.CountRow.ToString)
end;

procedure TFormSampleNoPackage.BitBtn3Click(Sender: TObject);
begin
  FOfficeWriter.URlFile := edtArqWriter.Text;
  FOfficeWriter.startDoc;
end;

procedure TFormSampleNoPackage.BitBtn4Click(Sender: TObject);
begin
  FOfficeWriter.CloseFile;
end;

procedure TFormSampleNoPackage.BitBtn5Click(Sender: TObject);
begin
  FOfficeWriter.saveFile(edtSalvar.Text);
end;

procedure TFormSampleNoPackage.BitBtn6Click(Sender: TObject);
begin
  FOfficeWriter.URlFile := edtArq.Text;
  FOfficeWriter.startDoc;
end;

procedure TFormSampleNoPackage.BitBtn7Click(Sender: TObject);
begin
  FOfficeWriter.
    setValue(mmo.Text).
    setBold(CB_Bold.Checked).
    setColorText(TOpenColor(cbColorWriter.Items.Objects[cbColorWriter.ItemIndex])).
    setFontHeight(StrToInt(edtFontHg.Text)).
    setUnderline(cbUnderline_wrt.Checked).
    setFontName(cbFontes.Text);
end;

procedure TFormSampleNoPackage.BitBtn8Click(Sender: TObject);
begin
  FOfficeWriter.setBold(true)
    .setFontHeight(16)
    .setValue('Título: Apresentando o componente Libre Office writer via Delphi <3' + #13#13);
  FOfficeWriter.gotoEndOfSentence;

  FOfficeWriter.setBold(false)
    .setFontHeight(12)
    .setValue('Neste exemplo estou mostrando a criação de documentos via código, de um jeito simples, rápido e fácil.' + #13);
  FOfficeWriter.gotoEndOfSentence;

  FOfficeWriter.setBold(false)
    .setFontHeight(12)
    .setValue('Espero que seja útil e que estejam gostando, lembrando o componente é open source e totalmente free!' + #13#13);
  FOfficeWriter.gotoEndOfSentence;

  FOfficeWriter.setBold(false)
    .setFontHeight(12)
    .SetBold(true)
    .setValue('Obrigado a todos pela presença!!!' + #13);
  FOfficeWriter.gotoEndOfSentence;
end;

procedure TFormSampleNoPackage.Button10Click(Sender: TObject);
var
  tpGrafico: TTypeChart;
begin
  if RBDefault.Checked then
    tpGrafico := ctDefault
  else if RBLine.Checked then
    tpGrafico := ctLine
  else if RBPie.Checked then
    tpGrafico := ctPie
  else if RBVertical.Checked then
    tpGrafico := ctVertical
  else
    tpGrafico := ctDefault;

   //Configure the chart settings
  FSettingsChart.Height := 11000;
  FSettingsChart.Width := 22000;
  FSettingsChart.Position_X := 1500;
  FSettingsChart.Position_Y := 12000;
  FSettingsChart.StartRow := StrToInt(edtLde.Text);
  FSettingsChart.EndRow := StrToInt(edtLAte.Text);
  FSettingsChart.PositionSheet := strToInt(edtPos.Text);
  FSettingsChart.StartColumn := uppercase(edtCde.Text);
  FSettingsChart.EndColumn := UpperCase(edtCAte.Text);
  FSettingsChart.ChartName := edtNomeGrafico.Text;
  FSettingsChart.typeChart := tpGrafico;

  FOfficeCalc.addChart(FSettingsChart);
end;

procedure TFormSampleNoPackage.Button11Click(Sender: TObject);
begin
  FOfficeCalc.CallConversorPDFTOSheet;
end;

procedure TFormSampleNoPackage.Button12Click(Sender: TObject);
begin
  if edtAba.Text = '' then
    edtAba.Text := 'Planilha 1';

  FOfficeCalc.SheetToDataSet(edtAba.Text);
end;

procedure TFormSampleNoPackage.CreateDemoSheet;
begin
  FOfficeCalc.DocVisible := CheckBox1.Checked;
  FOfficeCalc.startSheet;

  FOfficeCalc.SetValue(1, 'A', 'STATUS')
    .SetBorder([bAll], opBrown)
    .changeJustify(fthRIGHT, ftvTOP)
    .setBold(true)
    .changeFont('Arial', 12)
    .SetUnderline(true)
    .setColor(opWhite, opMagenta);

  FOfficeCalc.SetValue(1, 'B', 'VALOR').changeJustify(fthRIGHT, ftvTOP)
    .SetBorder([bAll], opBrown)
    .setBold(true)
    .changeFont('Arial', 12)
    .SetUnderline(true)
    .setColor(opWhite, opMagenta);

  FOfficeCalc.SetValue(2, 'B', 109, ftNumeric).SetBorder([bAll], opBrown);
  FOfficeCalc.SetValue(2, 'A', 'AGUA').SetBorder([bAll], opBrown);

  FOfficeCalc.SetValue(3, 'B', 105.55, ftNumeric).SetBorder([bAll], opBrown);
  FOfficeCalc.SetValue(3, 'A', 'LUZ').SetBorder([bAll], opBrown);

  FOfficeCalc.SetValue(4, 'B', 1005.22, ftNumeric);
  FOfficeCalc.SetValue(4, 'A', 'ALUGUEL');

  FOfficeCalc.SetValue(6, 'A', 'Total de linhas');
  FOfficeCalc.SetValue(6, 'B', FOfficeCalc.CountRow, ftNumeric);

  FOfficeCalc.SetValue(7, 'A', 'Total de Colunas');
  FOfficeCalc.SetValue(7, 'B', FOfficeCalc.CountCell, ftNumeric);

  FOfficeCalc.addNewSheet('A Receber', 1);

  FOfficeCalc.SetValue(1, 'A', 'VALOR')
    .SetBorder([bAll], opBrown)
    .changeJustify(fthRIGHT, ftvTOP)
    .setBold(true);

  FOfficeCalc.SetValue(1, 'B', 'DESC')
    .SetBorder([bAll], opBrown)
    .changeJustify(fthRIGHT, ftvTOP)
    .setBold(true)
    .changeFont('Arial', 12)
    .SetUnderline(true)
    .setColor(opWhite, opCiano);

  FOfficeCalc.SetValue(1, 'C', 'SOMA')
    .SetBorder([bAll], opBrown)
    .changeJustify(fthRIGHT, ftvTOP)
    .setBold(true)
    .changeFont('Arial', 12)
    .SetUnderline(true)
    .setColor(opWhite, opSoftRed);

  FOfficeCalc.SetValue(1, 'H', 'SOMA')
    .SetBorder([bAll], opBrown)
    .changeJustify(fthRIGHT, ftvTOP)
    .setBold(true)
    .changeFont('Arial', 12)
    .SetUnderline(true)
    .setColor(opWhite, opSoftRed);

  FOfficeCalc.SetValue(2, 'A', 200, ftNumeric);
  FOfficeCalc.SetValue(2, 'B', 'Emprestimo');
  FOfficeCalc.SetValue(2, 'C', 0, ftNumeric);

  FOfficeCalc.SetValue(3, 'A', 369.55, ftNumeric);
  FOfficeCalc.SetValue(3, 'B', 'Dividendos');
  FOfficeCalc.SetValue(3, 'C', 0, ftNumeric);

  FOfficeCalc.SetValue(4, 'A', 1585.22, ftNumeric);
  FOfficeCalc.SetValue(4, 'B', 'ALUGUEL');
  FOfficeCalc.SetValue(4, 'C', 0, ftNumeric);

  FOfficeCalc.SetValue(8, 'A', 1585.22, ftNumeric);
  FOfficeCalc.SetValue(8, 'B', 'Renda extra');
  FOfficeCalc.SetValue(8, 'C', 0, ftNumeric);

  FOfficeCalc.SetValue(15, 'A', 1585.22, ftNumeric);
  FOfficeCalc.SetValue(15, 'B', 'ALUGUEL 2');
  FOfficeCalc.SetValue(15, 'C', 0, ftNumeric);

  FOfficeCalc.SetValue(17, 'A', 'Total de linhas');
  FOfficeCalc.SetValue(17, 'B', FOfficeCalc.CountRow, ftNumeric);

  FOfficeCalc.SetValue(19, 'A', 'Total de Colunas');
  FOfficeCalc.SetValue(19, 'B', FOfficeCalc.CountCell, ftNumeric);
  FOfficeCalc.setFormula(20, 'A', '=A2+A3+A4+A15')
    .setBold(true);

  FOfficeCalc.positionSheetByName('Planilha1');

  //Configure the chart settings
  FSettingsChart.Height := 11000;
  FSettingsChart.Width := 22000;
  FSettingsChart.Position_X := 1500;
  FSettingsChart.Position_Y := 5000;
  FSettingsChart.StartRow := 0;
  FSettingsChart.EndRow := 3;
  FSettingsChart.PositionSheet := 0; //first tab
  FSettingsChart.StartColumn := 'A';
  FSettingsChart.EndColumn := 'B';
  FSettingsChart.ChartName := 'TestChart';
  FSettingsChart.typeChart := ctDefault;

  FOfficeCalc.addChart(FSettingsChart);

  FSettingsChart.typeChart := ctVertical;
  FOfficeCalc.addChart(FSettingsChart);

  FSettingsChart.typeChart := ctPie;
  FOfficeCalc.addChart(FSettingsChart);

  FSettingsChart.typeChart := ctLine;
  FOfficeCalc.addChart(FSettingsChart);
end;

procedure TFormSampleNoPackage.Button1Click(Sender: TObject);
begin
  //Dica: para desenvolver é mais facil uitilizar a propriedade OpenOffice_calc1.DocVisible := true;
  //Após desenv, alterar para false; em false, ganha desempenho e segurança, poís não ha risco do cliente fechar a planilha e perder o ponteiro
  FOfficeCalc.ExeThread(CreateDemoSheet);
end;

procedure TFormSampleNoPackage.Button2Click(Sender: TObject);
begin
  try
    FOfficeCalc.print;
    ShowMessage('Impressão Finalizada!');
  except
    ShowMessage('Erro ao imprimir!');
  end;
end;

procedure TFormSampleNoPackage.Button3Click(Sender: TObject);
begin
  if edtSalvar.Text <> '' then
    FOfficeCalc.saveFile(edtSalvar.Text);
end;

procedure TFormSampleNoPackage.Button4Click(Sender: TObject);
begin
  FOfficeCalc.CloseFile;
end;

procedure TFormSampleNoPackage.Button5Click(Sender: TObject);
begin
  FOfficeCalc.URlFile := edtArq.Text;
  FOfficeCalc.SheetName := edtAba.Text;
  FOfficeCalc.startSheet;
end;

procedure TFormSampleNoPackage.Button6Click(Sender: TObject);
begin
   FOfficeCalc.positionSheetByIndex(StrToIntDef(edtAba.Text,0));
end;

procedure TFormSampleNoPackage.btAddCellValueClick(Sender: TObject);
var
  tp: TTypeValue;
  jusH: THoriJustify;
  jusV: TVertJustify;
begin

  if chNumeric.Checked then
    tp := TTypeValue.ftNumeric
  else
    tp := TTypeValue.ftString;

  if RBhCenter.Checked then
    jusH := fthCENTER
  else if RBhLeft.Checked then
    jusH := fthLEFT
  else if RBhRight.Checked then
    jusH := fthRIGHT
  else
    jusH := fthSTANDARD;

  if RBvTop.Checked then
    jusV := ftvTOP
  else if RBvBottom.Checked then
    jusV := ftvBOTTOM
  else if RBvCenter.Checked then
    jusV := ftvCENTER
  else
    jusV := ftvSTANDARD;

  FOfficeCalc.SetValue(StrToInt(edtLinha.Text), edtColuna.Text, edtValor.Text, tp, cbQuebraLinha.Checked)
    .setBold(CBBold.Checked)
    .SetUnderline(CBUnderline.Checked)
    .setColor(TOpenColor(CBCorFont.Items.Objects[CBCorFont.ItemIndex]), TOpenColor(CBCorFundo.Items.Objects[CBCorFundo.ItemIndex]))
    .changeFont(cbFontes.Text, strToInt(edtTamanhoFonte.Text))
    .changeJustify(jusH, jusV)
    .SetBorder([bAll], opBlack)
    .setCellWidth(strToInt(edtWidth.Text));
end;

procedure TFormSampleNoPackage.btGetCellValueClick(Sender: TObject);
begin
  edtValor.Text := FOfficeCalc.GetValue(StrToInt(edtLinha.Text), edtColuna.Text).Value;
end;

procedure TFormSampleNoPackage.Button8Click(Sender: TObject);
begin
  FOfficeCalc.addNewSheet(edtAba.Text, StrToInt(edtPos.Text));
end;

procedure TFormSampleNoPackage.Button9Click(Sender: TObject);
begin
  FOfficeCalc.positionSheetByName(edtAba.Text);
end;

procedure TFormSampleNoPackage.FormCreate(Sender: TObject);
begin
  LoadAuxiliarItems;

  edtSalvar.Text := ExtractFileDir(ParamStr(0)) + '\SheetTest.ods';
  edtArq.Text := edtSalvar.Text;

  FOfficeCalc := TOpenOffice_calc.Create(self);
  FOfficeWriter := TOpenOffice_writer.Create(self);

  LoadSampleDataSetValues;
end;

procedure TFormSampleNoPackage.OpenOffice1AfterCloseSheet(sender: TObject);
begin
 // ShowMessage('Depois de Fechar');
end;

procedure TFormSampleNoPackage.OpenOffice1AfterGetValue(sender: TObject);
begin
 //ShowMessage('Depois de pegar o valor');
end;

procedure TFormSampleNoPackage.OpenOffice1AfterSetValue(sender: TObject);
begin
 // ShowMessage('Depois de Setar o valor');
end;

procedure TFormSampleNoPackage.OpenOffice1AfterStartSheet(sender: TObject);
begin
   // ShowMessage('iniciado');
end;

procedure TFormSampleNoPackage.OpenOffice1BeforeCloseSheet(sender: TObject);
begin
//    ShowMessage('Antes de fechar');
end;

procedure TFormSampleNoPackage.OpenOffice1BeforeGetValue(sender: TObject);
begin
  //ShowMessage('Antes de pegar o valor');
end;

procedure TFormSampleNoPackage.OpenOffice1BeforePrint(sender: TObject; var SetPrinter: TSetPrinter);
begin
  ShowMessage('Iniciando impressão...');
  SetPrinter.PaperSize_Width := 50000;
  SetPrinter.PaperSize_Height := 50000;
  SetPrinter.PrinterName := 'Microsoft Print to PDF';
  SetPrinter.Pages := '1; 3;';
end;

procedure TFormSampleNoPackage.OpenOffice1BeforeSetValue(sender: TObject);
begin
 //   ShowMessage('Antes de Setar o valor');
end;

procedure TFormSampleNoPackage.OpenOffice1BeforeStartSheet(sender: TObject);
begin
//  ShowMessage(  'Iniciando...'  );
end;

procedure TFormSampleNoPackage.TabSheet1Show(Sender: TObject);
begin
  cbFontes.Parent := TabSheet1;
  label4.Parent := tabSheet1;

  if FFontTop > 0 then
  begin
    cbFontes.top := FFontTop;
    cbFontes.Left := FFontLeft;
  end;

  label4.left := cbFontes.Left;
  label4.top := cbFontes.top - 15;
end;

procedure TFormSampleNoPackage.TabSheet2Show(Sender: TObject);
begin
  cbFontes.Parent := TabSheet2;
  label4.Parent := tabSheet2;

  FFontTop := cbFontes.top;
  FFontLeft := cbFontes.Left;

  cbFontes.top := 85;
  cbFontes.Left := 48;

  label4.left := cbFontes.Left;
  label4.top := cbFontes.top - 15;
end;

function EnumFontsProc(var LogFont: TLogFont; var TextMetric: TTextMetric; FontType: Integer; Data: Pointer): Integer; stdcall;
begin
  TStrings(Data).Add(LogFont.lfFaceName);
  Result := 1;
end;

procedure TFormSampleNoPackage.CapturarNomesDeFontes;
var
  DC: HDC;
begin

  DC := GetDC(0);

  EnumFonts(DC, nil, @EnumFontsProc, Pointer(cbFontes.Items));

  ReleaseDC(0, DC);

  cbFontes.Sorted := true;

  cbFontes.ItemIndex := 0;
end;

procedure TFormSampleNoPackage.CheckBox1Click(Sender: TObject);
begin
  if CheckBox1.Checked then
    lbl1.Caption := ' Para visualizar o arquivo gerado em modo visibilidade oculta, salvar o arquivo'
  else
    lbl1.Caption := '';
end;

procedure TFormSampleNoPackage.Exportarplanilha1Click(Sender: TObject);
begin
  FOfficeCalc.DocVisible := CheckBox1.Checked;
  FOfficeCalc.startSheet;
  FOfficeCalc.DataSetToSheet(cdsSampleValues);
end;

procedure TFormSampleNoPackage.LoadAuxiliarItems;
begin
  CBCorFont.Items.AddObject('Preto', TObject(opBlack));
  CBCorFont.Items.AddObject('Azul', TObject(opBlue));
  CBCorFont.Items.AddObject('Verde', TObject(opGreen));
  CBCorFont.Items.AddObject('Turquesa', TObject(optTurquesa));
  CBCorFont.Items.AddObject('Vermelho', TObject(opRed));
  CBCorFont.Items.AddObject('Magenta', TObject(opMagenta));
  CBCorFont.Items.AddObject('Marrom', TObject(opBrown));
  CBCorFont.Items.AddObject('Cinza0', TObject(opGray));
  CBCorFont.Items.AddObject('Cinza claro', TObject(opSoftGray));
  CBCorFont.Items.AddObject('Azul claro', TObject(opSoftBlue));
  CBCorFont.Items.AddObject('Cinza 6', TObject(opGreen6));
  CBCorFont.Items.AddObject('Ciano', TObject(opCiano));
  CBCorFont.Items.AddObject('Vermelho claro', TObject(opSoftRed));
  CBCorFont.Items.AddObject('Magenta claro', TObject(opSoftMagenta));
  CBCorFont.Items.AddObject('Amarelo', TObject(opYellow));
  CBCorFont.Items.AddObject('Branco', TObject(opWhite));
  CBCorFont.Items.AddObject('Cinza 30', TObject(opGray30));
  CBCorFont.Items.AddObject('Salmão', TObject(opSalmon));
  CBCorFont.Items.AddObject('Laranja', TObject(opOrange));
  CBCorFont.Items.AddObject('Laranja 80', TObject(opOrange80));
  CBCorFont.Items.AddObject('Bordo', TObject(opBordo));

  CBCorFundo.Items := CBCorFont.Items;
  cbColorWriter.Items := CBCorFont.Items;

  cbColorWriter.ItemIndex := 0;
  CBCorFundo.ItemIndex := 15;
  CBCorFont.ItemIndex := 0;

  edtTamanhoFonte.text := '10';

  FFontTop := 0;
  FFontLeft := 0;

  CapturarNomesDeFontes;
end;

procedure TFormSampleNoPackage.LoadSampleDataSetValues;
begin
  cdsSampleValues.CreateDataSet;

  cdsSampleValues.Insert;
  cdsSampleValuesID.Value := 1;
  cdsSampleValuesNome.Value := 'Daniel Fernandes';
  cdsSampleValuesIdade.Value := 28;
  cdsSampleValues.Post;

  cdsSampleValues.Insert;
  cdsSampleValuesID.Value := 2;
  cdsSampleValuesNome.Value := 'Bruna Silva';
  cdsSampleValuesIdade.Value := 29;
  cdsSampleValues.Post;

  cdsSampleValues.Insert;
  cdsSampleValuesID.Value := 3;
  cdsSampleValuesNome.Value := 'João do vale';
  cdsSampleValuesIdade.Value := 40;
  cdsSampleValues.Post;

  cdsSampleValues.Insert;
  cdsSampleValuesID.Value := 4;
  cdsSampleValuesNome.Value := 'Peter Parker';
  cdsSampleValuesIdade.Value := 25;
  cdsSampleValues.Post;
end;

end.

