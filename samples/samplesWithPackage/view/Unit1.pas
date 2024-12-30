unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, math, Vcl.Dialogs, Vcl.StdCtrls,  Vcl.ExtCtrls, Vcl.Buttons, Vcl.Mask,
  Vcl.ComCtrls, Data.DB, Datasnap.DBClient, Vcl.Grids, Vcl.DBGrids, Vcl.Menus, Datasnap.Provider,

  uOpenOffice.Calc,
  UOpenOffice.Writer,
  uOpenOffice.Collors,
  uOpenOffice.Helpers,
  uOpenOffice.SetPrinter,
  uOpenOffice;

type
  TForm1 = class(TForm)
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
    Button6: TButton;
    Button7: TButton;
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
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
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
    ClientDataSet1: TClientDataSet;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    ClientDataSet1ID: TIntegerField;
    ClientDataSet1Nome: TStringField;
    ClientDataSet1Idade: TIntegerField;
    PopupMenu1: TPopupMenu;
    Exportarplanilha1: TMenuItem;
    DataSetProvider1: TDataSetProvider;
    edtArqWriter: TLabeledEdit;
    OpenOffice_calc1: TOpenOffice_calc;
    OpenOffice_writer1: TOpenOffice_writer;
    Button11: TButton;
    Button12: TButton;
    lbl1: TLabel;
    lbl2: TLabel;
    CheckBox1: TCheckBox;
    edtWidth: TLabeledEdit;
    Button13: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
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
    procedure Button13Click(Sender: TObject);
  private
    FontTop, fontLeft : integer;
    SettingsChart: TSettingsChart;
    procedure CapturarNomesDeFontes;
    procedure CreateDemoSheet;

    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}


procedure TForm1.BitBtn10Click(Sender: TObject);
begin
   Openoffice_writer1.gotoEndOfSentence;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
  ShowMessage(Openoffice_calc1.CountCell.ToString);
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
  ShowMessage(Openoffice_calc1.CountRow.ToString)
end;

procedure TForm1.BitBtn3Click(Sender: TObject);
begin
   OpenOffice_writer1.DocVisible := True;
   Openoffice_writer1.URlFile   := edtArqWriter.Text;
   Openoffice_writer1.startDoc;
end;

procedure TForm1.BitBtn4Click(Sender: TObject);
begin
  Openoffice_writer1.CloseFile;
end;

procedure TForm1.BitBtn5Click(Sender: TObject);
begin
  Openoffice_writer1.SaveFile(edtSalvar.Text);
end;

procedure TForm1.BitBtn6Click(Sender: TObject);
begin
   Openoffice_writer1.URlFile   := edtArq.Text;
   Openoffice_writer1.startDoc;
end;

procedure TForm1.BitBtn7Click(Sender: TObject);
begin
  Openoffice_writer1.
  setValue(mmo.Text).
  SetBold(CB_Bold.Checked).
  setColorText(TOpenColor( cbColorWriter.Items.Objects[ cbColorWriter.ItemIndex])).
  setFontHeight(StrToInt(edtFontHg.Text)).
  SetUnderline(cbUnderline_wrt.Checked).
  setFontName(cbFontes.Text);
end;

procedure TForm1.BitBtn8Click(Sender: TObject);
begin
  Openoffice_writer1.SetBold(True)
     .setFontHeight(16)
     .setValue('Título: Apresentando o componente Libre Office writer via Delphi <3'+#13#13);
     Openoffice_writer1.gotoEndOfSentence;

  Openoffice_writer1.SetBold(false)
     .setFontHeight(12)
     .setValue('Neste exemplo estou mostrando a criação de documentos via código, de um jeito simples, rápido e fácil.'+#13 );
     Openoffice_writer1.gotoEndOfSentence;

     Openoffice_writer1.SetBold(false)
     .setFontHeight(12)
     .setValue('Espero que seja útil e que estejam gostando, lembrando o componente é open source e totalmente free!'+#13#13);
     Openoffice_writer1.gotoEndOfSentence;

     Openoffice_writer1.SetBold(false)
     .setFontHeight(12)
     .SetBold(True)
     .setValue('Obrigado a todos pela presença!!!' +#13);
     Openoffice_writer1.gotoEndOfSentence;
end;

procedure TForm1.Button10Click(Sender: TObject);
var tpGrafico : TTypeChart;
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
    SettingsChart.Height := 11000;
    SettingsChart.Width  := 22000;
    SettingsChart.Position_X := 1500;
    SettingsChart.Position_Y := 12000;
    SettingsChart.StartRow := StrToInt(edtLde.Text);
    SettingsChart.EndRow := StrToInt(edtLAte.Text);
    SettingsChart.PositionSheet := strToInt(edtPos.Text);
    SettingsChart.StartColumn := uppercase(edtCde.Text);
    SettingsChart.EndColumn := UpperCase(edtCAte.Text);
    SettingsChart.ChartName := edtNomeGrafico.Text;
    SettingsChart.TypeChart := tpGrafico;

   Openoffice_calc1.AddChart(SettingsChart);
end;

procedure TForm1.Button11Click(Sender: TObject);
begin
  OpenOffice_calc1.CallConversorPDFTOSheet;
end;

procedure TForm1.Button12Click(Sender: TObject);
begin
  if edtAba.Text = '' then
    edtAba.Text := 'Planilha 1';

  OpenOffice_calc1.SheetToDataSet(edtAba.Text);
end;

procedure TForm1.Button13Click(Sender: TObject);
begin
  OpenOffice_calc1.positionSheetByIndex(StrToIntDef(edtPos.Text,0));
end;

procedure TForm1.CreateDemoSheet;
begin
  OpenOffice_calc1.DocVisible := CheckBox1.Checked;
  Openoffice_calc1.StartSheet;

    Openoffice_calc1.SetValue(1,'A', 'STATUS')
       .SetBorder([bAll], opBrown)
       .ChangeJustify(  fthRIGHT , ftvTOP)
       .SetBold(True)
       .ChangeFont('Arial',12)
       .SetUnderline(True)
       .SetColor( opWhite ,opMagenta);

    Openoffice_calc1.SetValue(1,'B', 'VALOR').ChangeJustify(  fthRIGHT , ftvTOP)
       .SetBorder([bAll], opBrown)
       .SetBold(True)
       .ChangeFont('Arial',12)
       .SetUnderline(True)
       .SetColor( opWhite ,opMagenta);

    Openoffice_calc1.SetValue(2,'B', 109, ftNumeric).SetBorder([bAll], opBrown);
    Openoffice_calc1.SetValue(2,'A', 'AGUA').SetBorder([bAll], opBrown);

    Openoffice_calc1.SetValue(3,'B', 105.55, ftNumeric).SetBorder([bAll], opBrown);
    Openoffice_calc1.SetValue(3,'A', 'LUZ').SetBorder([bAll], opBrown);

    Openoffice_calc1.SetValue(4,'B', 1005.22, ftNumeric);
    Openoffice_calc1.SetValue(4,'A', 'ALUGUEL');

     Openoffice_calc1.SetValue(6,'A', 'Total de linhas');
     Openoffice_calc1.SetValue(6,'B', Openoffice_calc1.CountRow, ftNumeric);

    Openoffice_calc1.SetValue(7,'A', 'Total de Colunas');
    Openoffice_calc1.SetValue(7,'B', Openoffice_calc1.CountCell, ftNumeric);


    Openoffice_calc1.AddNewSheet('A Receber',1);

    Openoffice_calc1.SetValue(1,'A', 'VALOR')
       .SetBorder([bAll], opBrown)
       .ChangeJustify(  fthRIGHT , ftvTOP)
       .SetBold(True);

    Openoffice_calc1.SetValue(1,'B', 'DESC')
       .SetBorder([bAll], opBrown)
       .ChangeJustify(  fthRIGHT , ftvTOP)
       .SetBold(True)
       .ChangeFont('Arial',12)
       .SetUnderline(True)
       .SetColor( opWhite ,opCiano);

    Openoffice_calc1.SetValue(1,'C', 'SOMA')
       .SetBorder([bAll], opBrown)
       .ChangeJustify(  fthRIGHT , ftvTOP)
       .SetBold(True)
       .ChangeFont('Arial',12)
       .SetUnderline(True)
       .SetColor( opWhite ,opSoftRed);

    Openoffice_calc1.SetValue(1,'H', 'SOMA')
       .SetBorder([bAll], opBrown)
       .ChangeJustify(  fthRIGHT , ftvTOP)
       .SetBold(True)
       .ChangeFont('Arial',12)
       .SetUnderline(True)
       .SetColor( opWhite ,opSoftRed);

    Openoffice_calc1.SetValue(2,'A', 200, ftNumeric);
    Openoffice_calc1.SetValue(2,'B', 'Emprestimo');
    Openoffice_calc1.SetValue(2,'C', 0, ftNumeric);

    Openoffice_calc1.SetValue(3,'A', 369.55, ftNumeric);
    Openoffice_calc1.SetValue(3,'B', 'Dividendos');
    Openoffice_calc1.SetValue(3,'C', 0, ftNumeric);

    Openoffice_calc1.SetValue(4,'A', 1585.22, ftNumeric);
    Openoffice_calc1.SetValue(4,'B', 'ALUGUEL');
    Openoffice_calc1.SetValue(4,'C', 0, ftNumeric);

    Openoffice_calc1.SetValue(8,'A', 1585.22, ftNumeric);
    Openoffice_calc1.SetValue(8,'B', 'Renda extra');
    Openoffice_calc1.SetValue(8,'C', 0, ftNumeric);

    Openoffice_calc1.SetValue(15,'A', 1585.22, ftNumeric);
    Openoffice_calc1.SetValue(15,'B', 'ALUGUEL 2');
    Openoffice_calc1.SetValue(15,'C', 0, ftNumeric);

    Openoffice_calc1.SetValue(17,'A', 'Total de linhas');
    Openoffice_calc1.SetValue(17,'B', Openoffice_calc1.CountRow, ftNumeric);

    Openoffice_calc1.SetValue(19,'A', 'Total de Colunas');
    Openoffice_calc1.SetValue(19,'B', Openoffice_calc1.CountCell, ftNumeric);
    Openoffice_calc1.SetFormula(20,'A',  '=A2+A3+A4+A15')
      .SetBold(True);


    Openoffice_calc1.PositionSheetByName('Planilha1');

    //Configure the chart settings
    SettingsChart.Height := 11000;
    SettingsChart.Width  := 22000;
    SettingsChart.Position_X := 1500;
    SettingsChart.Position_Y := 5000;
    SettingsChart.StartRow := 0;
    SettingsChart.EndRow := 3;
    SettingsChart.PositionSheet := 0; //first tab
    SettingsChart.StartColumn := 'A';
    SettingsChart.EndColumn := 'B';
    SettingsChart.ChartName := 'TestChart';
    SettingsChart.TypeChart := ctDefault;

    Openoffice_calc1.AddChart(SettingsChart);

    SettingsChart.TypeChart := ctVertical;
    Openoffice_calc1.AddChart(SettingsChart);

    SettingsChart.TypeChart := ctPie;
    Openoffice_calc1.AddChart(SettingsChart);

    SettingsChart.TypeChart := ctLine;
    Openoffice_calc1.AddChart(SettingsChart);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  //Dica: para desenvolver é mais facil uitilizar a propriedade OpenOffice_calc1.DocVisible := True;
  //Após desenv, alterar para false; em false, ganha desempenho e segurança, poís não ha risco do cliente fechar a planilha e perder o ponteiro
  Openoffice_calc1.ExeThread(CreateDemoSheet);
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
   try
      Openoffice_calc1.Print;
      ShowMessage('Impressão Finalizada!');
   except
      ShowMessage('Erro ao imprimir!');
   end;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  if edtSalvar.Text <> '' then
    Openoffice_calc1.SaveFile(edtSalvar.Text);
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
  Openoffice_calc1.CloseFile;
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
   Openoffice_calc1.URlFile   := edtArq.Text;
   Openoffice_calc1.SheetName := edtAba.Text;
   Openoffice_calc1.StartSheet;
end;

procedure TForm1.Button6Click(Sender: TObject);
var tp   : TTypeValue;
    jusH : THoriJustify;
    jusV : TVertJustify;
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

  Openoffice_calc1.SetValue(StrToInt(edtLinha.Text), edtColuna.Text, edtValor.Text, tp,cbQuebraLinha.Checked)
    .SetBold(CBBold.Checked)
    .SetUnderline(CBUnderline.Checked)
    .SetColor(TOpenColor( CBCorFont.Items.Objects[ CBCorFont.ItemIndex]),TOpenColor( CBCorFundo.Items.Objects[ CBCorFundo.ItemIndex]))
    .ChangeFont(cbFontes.Text,strToInt(edtTamanhoFonte.Text))
    .ChangeJustify(jusH,jusV)
    .SetBorder([bAll], opBlack)
    .SetCellWidth(strToInt(edtWidth.Text));
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
  edtValor.Text := Openoffice_calc1.GetValue( StrToInt(edtLinha.Text), edtColuna.Text).Value;
end;

procedure TForm1.Button8Click(Sender: TObject);
begin
  Openoffice_calc1.AddNewSheet(edtAba.Text, StrToInt(edtPos.Text));
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
  Openoffice_calc1.PositionSheetByName(edtAba.Text);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  CBCorFont.Items.AddObject('Preto', TObject(opBlack));
  CBCorFont.Items.AddObject('Azul'    , TObject(opBlue));
  CBCorFont.Items.AddObject('Verde'   , TObject(opGreen));
  CBCorFont.Items.AddObject('Turquesa', TObject(optTurquesa));
  CBCorFont.Items.AddObject('Vermelho', TObject(opRed));
  CBCorFont.Items.AddObject('Magenta' , TObject(opMagenta));
  CBCorFont.Items.AddObject('Marrom'  , TObject(opBrown));
  CBCorFont.Items.AddObject('Cinza0', TObject(opGray));
  CBCorFont.Items.AddObject('Cinza claro', TObject(opSoftGray));
  CBCorFont.Items.AddObject('Azul claro', TObject(opSoftBlue));
  CBCorFont.Items.AddObject('Cinza 6',TObject(opGreen6));
  CBCorFont.Items.AddObject('Ciano', TObject(opCiano));
  CBCorFont.Items.AddObject('Vermelho claro', TObject(opSoftRed));
  CBCorFont.Items.AddObject('Magenta claro', TObject(opSoftMagenta));
  CBCorFont.Items.AddObject('Amarelo', TObject(opYellow));
  CBCorFont.Items.AddObject('Branco' , TObject(opWhite));
  CBCorFont.Items.AddObject('Cinza 30', TObject(opGray30));
  CBCorFont.Items.AddObject('Salmão', TObject(opSalmon));
  CBCorFont.Items.AddObject('Laranja', TObject(opOrange));
  CBCorFont.Items.AddObject('Laranja 80', TObject(opOrange80));
  CBCorFont.Items.AddObject('Bordo',TObject(opBordo));

  CBCorFundo.Items    :=  CBCorFont.Items;
  cbColorWriter.Items := CBCorFont.Items;

  cbColorWriter.ItemIndex :=0;
  CBCorFundo.ItemIndex := 15;
  CBCorFont.ItemIndex  := 0;

  edtTamanhoFonte.text := '10';

  fontTop  := 0;
  fontLeft := 0;

  CapturarNomesDeFontes;

  edtSalvar.Text    := ExtractFileDir(GetCurrentDir)+'\';
  edtArq.Text       := edtSalvar.Text;

  ClientDataSet1.Open;
  ClientDataSet1.Insert;
  ClientDataSet1ID.Value := 1;
  ClientDataSet1Nome.Value := 'Daniel Fernandes';
  ClientDataSet1Idade.Value := 28;
  ClientDataSet1.Post;

  ClientDataSet1.Insert;
  ClientDataSet1ID.Value := 2;
  ClientDataSet1Nome.Value := 'Bruna Silva';
  ClientDataSet1Idade.Value := 29;
  ClientDataSet1.Post;

  ClientDataSet1.Insert;
  ClientDataSet1ID.Value := 3;
  ClientDataSet1Nome.Value := 'João do vale';
  ClientDataSet1Idade.Value := 40;
  ClientDataSet1.Post;

  ClientDataSet1.Insert;
  ClientDataSet1ID.Value := 4;
  ClientDataSet1Nome.Value := 'Peter Parker';
  ClientDataSet1Idade.Value := 25;
  ClientDataSet1.Post;

end;


procedure TForm1.OpenOffice1AfterCloseSheet(sender: TObject);
begin
 // ShowMessage('Depois de Fechar');
end;

procedure TForm1.OpenOffice1AfterGetValue(sender: TObject);
begin
 //ShowMessage('Depois de pegar o valor');
end;

procedure TForm1.OpenOffice1AfterSetValue(sender: TObject);
begin
 // ShowMessage('Depois de Setar o valor');
end;

procedure TForm1.OpenOffice1AfterStartSheet(sender: TObject);
begin
   // ShowMessage('iniciado');
end;

procedure TForm1.OpenOffice1BeforeCloseSheet(sender: TObject);
begin
//    ShowMessage('Antes de fechar');
end;

procedure TForm1.OpenOffice1BeforeGetValue(sender: TObject);
begin
  //ShowMessage('Antes de pegar o valor');
end;

procedure TForm1.OpenOffice1BeforePrint(sender: TObject; var SetPrinter: TSetPrinter);
begin
  ShowMessage('Iniciando impressão...');
  SetPrinter.PaperSize_Width  := 50000;
  SetPrinter.PaperSize_Height := 50000;
  SetPrinter.PrinterName      := 'Microsoft Print to PDF';
  SetPrinter.Pages            := '1; 3;';
end;

procedure TForm1.OpenOffice1BeforeSetValue(sender: TObject);
begin
 //   ShowMessage('Antes de Setar o valor');
end;

procedure TForm1.OpenOffice1BeforeStartSheet(sender: TObject);
begin
//  ShowMessage(  'Iniciando...'  );
end;

procedure TForm1.TabSheet1Show(Sender: TObject);
begin
   cbFontes.Parent :=  TabSheet1;
   label4.Parent   :=  tabSheet1;

   if fontTop > 0 then
   begin
     cbFontes.top   := fontTop;
     cbFontes.Left  := fontLeft;
   end;

   label4.left    :=  cbFontes.Left;
   label4.top     :=  cbFontes.top -15;
end;

procedure TForm1.TabSheet2Show(Sender: TObject);
begin
   cbFontes.Parent :=  TabSheet2;
   label4.Parent   :=  tabSheet2;

   fontTop  := cbFontes.top;
   fontLeft := cbFontes.Left;

   cbFontes.top   := 85;
   cbFontes.Left  := 48;

   label4.left    := cbFontes.Left;
   label4.top     :=  cbFontes.top -15;
end;

function EnumFontsProc(var LogFont: TLogFont; var TextMetric: TTextMetric;
  FontType: Integer; Data: Pointer): Integer; stdcall;
begin
  TStrings(Data).Add(LogFont.lfFaceName);
  Result := 1;
end;

procedure TForm1.CapturarNomesDeFontes;
var
  DC: HDC;
begin

  DC := GetDC(0);

  EnumFonts(DC, nil, @EnumFontsProc, Pointer(cbFontes.Items));

  ReleaseDC(0, DC);

  cbFontes.Sorted := True;

  cbFontes.ItemIndex := 0;

end;

procedure TForm1.CheckBox1Click(Sender: TObject);
begin
  if CheckBox1.Checked then
    lbl1.Caption := ' Para visualizar o arquivo gerado em modo visibilidade oculta, salvar o arquivo'
  else
    lbl1.Caption := '';
end;

procedure TForm1.Exportarplanilha1Click(Sender: TObject);
begin
   OpenOffice_calc1.DocVisible := CheckBox1.Checked;
   Openoffice_calc1.StartSheet;
   Openoffice_calc1.DataSetToSheet(ClientDataSet1);
end;

end.
