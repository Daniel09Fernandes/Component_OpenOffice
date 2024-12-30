object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Demo'
  ClientHeight = 599
  ClientWidth = 665
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Position = poDesktopCenter
  OnCreate = FormCreate
  TextHeight = 13
  object PageControl2: TPageControl
    Left = 0
    Top = 0
    Width = 665
    Height = 599
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = 'Calc'
      OnShow = TabSheet1Show
      object Bevel1: TBevel
        Left = 0
        Top = 51
        Width = 399
        Height = 142
        Style = bsRaised
      end
      object Bevel2: TBevel
        Left = 0
        Top = 279
        Width = 636
        Height = 118
        Style = bsRaised
      end
      object Bevel3: TBevel
        Left = 405
        Top = 120
        Width = 231
        Height = 72
        Style = bsRaised
      end
      object Bevel4: TBevel
        Left = -270
        Top = -5
        Width = 669
        Height = 50
        Style = bsRaised
      end
      object Label1: TLabel
        Left = 6
        Top = 5
        Width = 71
        Height = 13
        Caption = 'tamanho fonte'
      end
      object Label2: TLabel
        Left = 6
        Top = 50
        Width = 61
        Height = 13
        Caption = 'Cor da fonte'
      end
      object Label3: TLabel
        Left = 163
        Top = 50
        Width = 63
        Height = 13
        Caption = 'Cor do fundo'
      end
      object Label4: TLabel
        Left = 93
        Top = 5
        Width = 33
        Height = 13
        Caption = 'Fontes'
      end
      object lbl1: TLabel
        Left = 3
        Top = 529
        Width = 3
        Height = 13
      end
      object lbl2: TLabel
        Left = 3
        Top = 548
        Width = 3
        Height = 13
      end
      object BitBtn1: TBitBtn
        Left = 180
        Top = 160
        Width = 62
        Height = 25
        Caption = 'Qde Coluna'
        TabOrder = 0
        OnClick = BitBtn1Click
      end
      object BitBtn2: TBitBtn
        Left = 252
        Top = 160
        Width = 56
        Height = 25
        Caption = 'Qde Linha'
        TabOrder = 1
        OnClick = BitBtn2Click
      end
      object Button1: TButton
        Left = 3
        Top = 363
        Width = 121
        Height = 30
        Caption = 'Criar nova planilha'
        TabOrder = 2
        OnClick = Button1Click
      end
      object Button10: TButton
        Left = 6
        Top = 235
        Width = 75
        Height = 21
        Caption = 'Add grafico'
        TabOrder = 3
        OnClick = Button10Click
      end
      object Button2: TButton
        Left = 263
        Top = 363
        Width = 81
        Height = 30
        Caption = 'Imprimir '
        TabOrder = 4
        OnClick = Button2Click
      end
      object Button3: TButton
        Left = 352
        Top = 363
        Width = 121
        Height = 30
        Caption = 'Salvar'
        TabOrder = 5
        OnClick = Button3Click
      end
      object Button4: TButton
        Left = 480
        Top = 363
        Width = 67
        Height = 30
        Caption = 'Fechar'
        TabOrder = 6
        OnClick = Button4Click
      end
      object Button5: TButton
        Left = 134
        Top = 363
        Width = 121
        Height = 30
        Caption = 'Carregar documeto'
        TabOrder = 7
        OnClick = Button5Click
      end
      object Button6: TButton
        Left = 84
        Top = 160
        Width = 89
        Height = 25
        Caption = 'Adicionar Valor'
        TabOrder = 8
        OnClick = Button6Click
      end
      object Button7: TButton
        Left = 6
        Top = 160
        Width = 75
        Height = 25
        Caption = 'Pegar valor'
        TabOrder = 9
        OnClick = Button7Click
      end
      object Button8: TButton
        Left = 413
        Top = 166
        Width = 57
        Height = 25
        Caption = 'Adicionar'
        TabOrder = 10
        OnClick = Button8Click
      end
      object Button9: TButton
        Left = 472
        Top = 166
        Width = 76
        Height = 25
        Caption = 'Trocar - Nome'
        TabOrder = 11
        OnClick = Button9Click
      end
      object CBBold: TCheckBox
        Left = 180
        Top = 97
        Width = 60
        Height = 17
        Caption = 'Negrito'
        TabOrder = 12
      end
      object CBCorFont: TComboBox
        Left = 6
        Top = 69
        Width = 145
        Height = 21
        TabOrder = 13
        Text = 'opBlack = 0,'
      end
      object CBCorFundo: TComboBox
        Left = 163
        Top = 69
        Width = 145
        Height = 21
        ItemIndex = 0
        TabOrder = 14
        Text = 'opBlack = 0,'
        Items.Strings = (
          'opBlack = 0,'
          'opBlue = 128,'
          'opGreen = 32768,'
          'optTurquesa = 32896,'
          'opRed = 8388608,'
          'opMagenta = 8388736,'
          'opBrown = 8421376,'
          'opGray = 8421504,'
          'opSoftGray = 12632256,'
          'opSoftBlue = 255,'
          'opGreen6 = 4057917,'
          'opCiano = 65535,'
          'opSoftRed = 16711680,'
          'opSoftMagenta = 16711935,'
          'opYellow = 16776960,'
          'opWhite = 16777215,'
          'opGray30 = 11776947,'
          'opSalmon = 26316,'
          'opOrange = 16750950,'
          'opOrange80 = 10066431,'
          'opBordo = 16777164')
      end
      object cbFontes: TComboBox
        Left = 93
        Top = 21
        Width = 294
        Height = 21
        TabOrder = 15
      end
      object cbQuebraLinha: TCheckBox
        Left = 6
        Top = 97
        Width = 97
        Height = 17
        Caption = 'Quebra de linha'
        TabOrder = 16
      end
      object CBUnderline: TCheckBox
        Left = 236
        Top = 97
        Width = 82
        Height = 17
        Caption = 'Sublinhado'
        TabOrder = 17
      end
      object chNumeric: TCheckBox
        Left = 103
        Top = 97
        Width = 76
        Height = 17
        Caption = 'is Numeric?'
        TabOrder = 18
      end
      object edtAba: TLabeledEdit
        Left = 453
        Top = 139
        Width = 170
        Height = 21
        EditLabel.Width = 73
        EditLabel.Height = 13
        EditLabel.Caption = 'Aba da planilha'
        TabOrder = 19
        Text = ''
      end
      object edtArq: TLabeledEdit
        Left = 3
        Top = 337
        Width = 537
        Height = 21
        EditLabel.Width = 99
        EditLabel.Height = 13
        EditLabel.Caption = 'Carregar documento'
        TabOrder = 20
        Text = ''
      end
      object edtCAte: TLabeledEdit
        Left = 70
        Top = 210
        Width = 53
        Height = 21
        EditLabel.Width = 53
        EditLabel.Height = 13
        EditLabel.Caption = 'Coluna Ate'
        TabOrder = 21
        Text = ''
      end
      object edtCde: TLabeledEdit
        Left = 6
        Top = 210
        Width = 52
        Height = 21
        EditLabel.Width = 48
        EditLabel.Height = 13
        EditLabel.Caption = 'Coluna de'
        TabOrder = 22
        Text = ''
      end
      object edtColuna: TLabeledEdit
        Left = 6
        Top = 133
        Width = 34
        Height = 21
        CharCase = ecUpperCase
        EditLabel.Width = 33
        EditLabel.Height = 13
        EditLabel.Caption = 'Coluna'
        TabOrder = 23
        Text = ''
      end
      object edtLAte: TLabeledEdit
        Left = 186
        Top = 210
        Width = 40
        Height = 21
        EditLabel.Width = 41
        EditLabel.Height = 13
        EditLabel.Caption = 'linha ate'
        NumbersOnly = True
        TabOrder = 24
        Text = ''
      end
      object edtLde: TLabeledEdit
        Left = 134
        Top = 210
        Width = 42
        Height = 21
        EditLabel.Width = 40
        EditLabel.Height = 13
        EditLabel.Caption = 'Linha de'
        NumbersOnly = True
        TabOrder = 25
        Text = ''
      end
      object edtLinha: TLabeledEdit
        Left = 46
        Top = 133
        Width = 35
        Height = 21
        EditLabel.Width = 25
        EditLabel.Height = 13
        EditLabel.Caption = 'Linha'
        NumbersOnly = True
        TabOrder = 26
        Text = ''
      end
      object edtNomeGrafico: TLabeledEdit
        Left = 236
        Top = 210
        Width = 400
        Height = 21
        EditLabel.Width = 64
        EditLabel.Height = 13
        EditLabel.Caption = 'Nome Grafico'
        TabOrder = 27
        Text = ''
      end
      object edtPos: TLabeledEdit
        Left = 413
        Top = 139
        Width = 34
        Height = 21
        EditLabel.Width = 36
        EditLabel.Height = 13
        EditLabel.Caption = 'Posi'#231#227'o'
        TabOrder = 28
        Text = ''
      end
      object edtSalvar: TLabeledEdit
        Left = 3
        Top = 296
        Width = 537
        Height = 21
        EditLabel.Width = 47
        EditLabel.Height = 13
        EditLabel.Caption = 'Salvar em'
        TabOrder = 29
        Text = 'c:\'
      end
      object edtValor: TLabeledEdit
        Left = 87
        Top = 132
        Width = 221
        Height = 21
        EditLabel.Width = 24
        EditLabel.Height = 13
        EditLabel.Caption = 'Valor'
        TabOrder = 30
        Text = ''
      end
      object Graficos: TGroupBox
        Left = 405
        Top = 44
        Width = 123
        Height = 70
        Caption = 'Tipos de Graficos'
        TabOrder = 31
        object RBDefault: TRadioButton
          Left = 9
          Top = 20
          Width = 56
          Height = 17
          Caption = 'Default'
          Checked = True
          TabOrder = 0
          TabStop = True
        end
        object RBVertical: TRadioButton
          Left = 9
          Top = 41
          Width = 48
          Height = 17
          Caption = 'Vertical'
          TabOrder = 1
        end
        object RBPie: TRadioButton
          Left = 74
          Top = 20
          Width = 116
          Height = 17
          Caption = 'Pie'
          TabOrder = 2
        end
        object RBLine: TRadioButton
          Left = 74
          Top = 41
          Width = 116
          Height = 17
          Caption = 'Line'
          TabOrder = 3
        end
      end
      object GroupBox1: TGroupBox
        Left = 405
        Top = 3
        Width = 231
        Height = 37
        Caption = 'Alinhamento horizontal'
        TabOrder = 33
        object RBhCenter: TRadioButton
          Left = 87
          Top = 18
          Width = 58
          Height = 17
          Caption = 'Centro'
          TabOrder = 0
        end
        object RBhLeft: TRadioButton
          Left = 3
          Top = 18
          Width = 70
          Height = 17
          Caption = 'Esquerda'
          TabOrder = 1
        end
        object RBhRight: TRadioButton
          Left = 160
          Top = 18
          Width = 113
          Height = 17
          Caption = 'Direita'
          TabOrder = 2
        end
        object PageControl1: TPageControl
          Left = 88
          Top = 40
          Width = 289
          Height = 193
          TabOrder = 3
        end
      end
      object GroupBox2: TGroupBox
        Left = 534
        Top = 44
        Width = 102
        Height = 70
        Caption = 'Alinha. Vertical'
        TabOrder = 32
        object RBvTop: TRadioButton
          Left = 32
          Top = 15
          Width = 113
          Height = 17
          Caption = 'Cima'
          TabOrder = 0
        end
        object RBvBottom: TRadioButton
          Left = 32
          Top = 51
          Width = 113
          Height = 17
          Caption = 'Baixo'
          TabOrder = 1
        end
        object RBvCenter: TRadioButton
          Left = 32
          Top = 33
          Width = 113
          Height = 17
          Caption = 'Centro'
          TabOrder = 2
        end
      end
      object edtTamanhoFonte: TEdit
        Left = 6
        Top = 21
        Width = 61
        Height = 21
        NumbersOnly = True
        TabOrder = 34
      end
      object DBGrid1: TDBGrid
        Left = 3
        Top = 403
        Width = 633
        Height = 120
        DataSource = DataSource1
        PopupMenu = PopupMenu1
        TabOrder = 35
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object Button11: TButton
        Left = 6
        Top = 256
        Width = 97
        Height = 25
        Caption = 'Pdf para planilha'
        TabOrder = 36
        OnClick = Button11Click
      end
      object Button12: TButton
        Left = 109
        Top = 256
        Width = 117
        Height = 25
        Caption = 'Planilha para DataSet'
        TabOrder = 37
        OnClick = Button12Click
      end
      object CheckBox1: TCheckBox
        Left = 236
        Top = 235
        Width = 181
        Height = 17
        Caption = 'Visualizar gera'#231#227'o do documento? '
        Checked = True
        State = cbChecked
        TabOrder = 38
        OnClick = CheckBox1Click
      end
      object edtWidth: TLabeledEdit
        Left = 314
        Top = 132
        Width = 44
        Height = 21
        CharCase = ecUpperCase
        EditLabel.Width = 48
        EditLabel.Height = 13
        EditLabel.Caption = 'Cel lWidth'
        EditLabel.Color = clBackground
        EditLabel.ParentColor = False
        ImeName = 'Portuguese (Brazilian ABNT)'
        TabOrder = 39
        Text = '5000'
      end
      object Button13: TButton
        Left = 551
        Top = 166
        Width = 73
        Height = 25
        Caption = 'Trocar - Index'
        TabOrder = 40
        OnClick = Button13Click
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Writer'
      ImageIndex = 1
      OnShow = TabSheet2Show
      object Label5: TLabel
        Left = 112
        Top = 29
        Width = 73
        Height = 13
        Caption = 'Tamanho fonte'
      end
      object Label6: TLabel
        Left = 215
        Top = 29
        Width = 17
        Height = 13
        Caption = 'Cor'
      end
      object BitBtn3: TBitBtn
        Left = 48
        Top = 112
        Width = 75
        Height = 25
        Caption = 'Iniciar'
        TabOrder = 0
        OnClick = BitBtn3Click
      end
      object BitBtn4: TBitBtn
        Left = 129
        Top = 112
        Width = 75
        Height = 25
        Caption = 'Fechar'
        TabOrder = 1
        OnClick = BitBtn4Click
      end
      object BitBtn5: TBitBtn
        Left = 224
        Top = 112
        Width = 75
        Height = 25
        Caption = 'Salvar'
        TabOrder = 2
        OnClick = BitBtn5Click
      end
      object BitBtn6: TBitBtn
        Left = 305
        Top = 112
        Width = 75
        Height = 25
        Caption = 'Carregar'
        TabOrder = 3
        OnClick = BitBtn6Click
      end
      object mmo: TMemo
        Left = 48
        Top = 143
        Width = 457
        Height = 210
        Lines.Strings = (
          'Ol'#225' mundo')
        TabOrder = 4
      end
      object BitBtn7: TBitBtn
        Left = 48
        Top = 359
        Width = 129
        Height = 25
        Caption = 'Adicionar texto'
        TabOrder = 5
        OnClick = BitBtn7Click
      end
      object edtFontHg: TEdit
        Left = 112
        Top = 48
        Width = 65
        Height = 21
        NumbersOnly = True
        TabOrder = 6
        Text = '10'
      end
      object CB_Bold: TCheckBox
        Left = 48
        Top = 29
        Width = 41
        Height = 17
        Caption = 'Bold'
        TabOrder = 7
      end
      object CBErase: TCheckBox
        Left = 48
        Top = 6
        Width = 161
        Height = 17
        Caption = 'Apar texto antes de escrever'
        Checked = True
        State = cbChecked
        TabOrder = 8
      end
      object BitBtn8: TBitBtn
        Left = 401
        Top = 112
        Width = 104
        Height = 25
        Caption = 'Montar exemplo'
        TabOrder = 9
        OnClick = BitBtn8Click
      end
      object cbUnderline_wrt: TCheckBox
        Left = 48
        Top = 52
        Width = 41
        Height = 17
        Caption = 'UnderLine'
        TabOrder = 10
      end
      object cbColorWriter: TComboBox
        Left = 215
        Top = 48
        Width = 145
        Height = 21
        ItemIndex = 0
        TabOrder = 11
        Text = 'opBlack = 0,'
        Items.Strings = (
          'opBlack = 0,'
          'opBlue = 128,'
          'opGreen = 32768,'
          'optTurquesa = 32896,'
          'opRed = 8388608,'
          'opMagenta = 8388736,'
          'opBrown = 8421376,'
          'opGray = 8421504,'
          'opSoftGray = 12632256,'
          'opSoftBlue = 255,'
          'opGreen6 = 4057917,'
          'opCiano = 65535,'
          'opSoftRed = 16711680,'
          'opSoftMagenta = 16711935,'
          'opYellow = 16776960,'
          'opWhite = 16777215,'
          'opGray30 = 11776947,'
          'opSalmon = 26316,'
          'opOrange = 16750950,'
          'opOrange80 = 10066431,'
          'opBordo = 16777164')
      end
      object BitBtn10: TBitBtn
        Left = 195
        Top = 359
        Width = 198
        Height = 25
        Caption = 'Mover cursor para o final da pagina'
        TabOrder = 12
        OnClick = BitBtn10Click
      end
      object edtArqWriter: TLabeledEdit
        Left = 48
        Top = 417
        Width = 457
        Height = 21
        EditLabel.Width = 99
        EditLabel.Height = 13
        EditLabel.Caption = 'Carregar documento'
        TabOrder = 13
        Text = ''
      end
    end
  end
  object ClientDataSet1: TClientDataSet
    PersistDataPacket.Data = {
      4C0000009619E0BD0100000018000000030000000000030000004C0002494404
      00010000000000044E6F6D650100490000000100055749445448020002001400
      05496461646504000100000000000000}
    Active = True
    Aggregates = <>
    Params = <>
    Left = 44
    Top = 464
    object ClientDataSet1ID: TIntegerField
      FieldName = 'ID'
    end
    object ClientDataSet1Nome: TStringField
      FieldName = 'Nome'
    end
    object ClientDataSet1Idade: TIntegerField
      FieldName = 'Idade'
    end
  end
  object DataSource1: TDataSource
    DataSet = ClientDataSet1
    Left = 124
    Top = 464
  end
  object PopupMenu1: TPopupMenu
    Left = 468
    Top = 480
    object Exportarplanilha1: TMenuItem
      Caption = 'Exportar planilha'
      OnClick = Exportarplanilha1Click
    end
  end
  object DataSetProvider1: TDataSetProvider
    UpdateMode = upWhereKeyOnly
    Left = 212
    Top = 464
  end
  object OpenOffice_calc1: TOpenOffice_Calc
    DocVisible = False
    Left = 296
    Top = 232
  end
  object OpenOffice_writer1: TOpenOffice_writer
    DocVisible = True
    Left = 420
    Top = 224
  end
end
