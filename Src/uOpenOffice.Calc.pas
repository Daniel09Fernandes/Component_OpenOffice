{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOffice_calc.pas }
{ Developer: Daniel Fernandes Rodrigures }
{ Email    : danielfernandesroddrigues@gmail.com }
{ this unit is a part of the Open Source. }
{ licensed under a MIT - see LICENSE.md}

{ ******************************************************* }

{ Documentation:                                           }
{
  https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Editing_Spreadsheet_Documents
  https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Cells_and_Ranges
  https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Templates
  https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/StarDesktop
}
{ ******************************************************* }

unit uOpenOffice.Calc;

interface

uses
  System.Classes, data.DB, ActiveX, uOpenOffice,
  dbWeb, ComObj, XMLDoc, XMLIntf, Vcl.Dialogs, System.Variants,
  Windows, uOpenOffice.Events, Datasnap.DBClient, System.SysUtils;

type
  TTypeValue = (ftString, ftNumeric);
  //0; //1234,2
  //1; //1234
  //3; //1.234
  //4; //1.234,20
  //10; //123420%
  //11; //123420,00%
  //20; //R$ 1.234
  //21; //R$ 1.234,20
  //24; // 1.234,20 BRL
  //30; // DD/MM/YY
  //31; // DDDD DD/MM/YY
  //32; // MM/YY
  //33; // DD/MMMM
  //34; // MMMM
  //35; // TRIMESTRE
  //36; // DD/MM/YYYY
  //39; // DD MMMM YY
  //40; // DD MMMM YYYY
  TNumberMask = (Default, NoDecimalSeparator, WithDecimalSeparator = 3, WithDecimalSeparatorAndComma, Percent = 10, PercentWithComma, CurrencySymbolWithoutDecimal = 20,
                   CurrencySymbolWithDecimal, CurrencySuffix = 24, Date_dd_mm_yy = 30, Date_dddd_dd_mm_yyy,  Date_mm_yy, Date_dd_mmmm, Date_mmmm,
                   Date_QUARTER, Date_Default, Date_dd_mmmm_yy = 39, Date_dd_mmmm_yyyy);


  TFieldsSheet = record
  private
  var
    ArrFields: array of string;
    
    procedure SetArrayFieldsSheet;
  public
    function GetField(AIndex: integer): String;
    function GetIndex(ANameField: String): Integer;
  end;

  TOpenOffice_Calc = class(TOpenOffice)
  private
    const
      DefaultNewSheetNamePT = 'Planilha1';
      DefaultNewSheetNameEn = 'Sheet1';
      procedure ValidateSheetName;
    var
      //--------events------//
      FOnBeforeStartFile: TBeforeStartFile;
      FOnAfterStartFile : TAfterStartFile;
      //--------------------//
      FFields: TFieldsSheet;
      FSheetName: string;
      FNumberMask: TNumberMask;
      FValue: string;

    procedure SetSheetName(const Value: string);
  published
    property ServicesManager: OleVariant read objServiceManager;
    property Cell: OleVariant read ObjCell write ObjCell;
    property OSCalc: OleVariant read ObjSCalc write ObjSCalc;
    property Fields: TFieldsSheet read FFields;
    property CoreReflection :OleVariant read ObjCoreReflection;
    property SheetName: string read FSheetName write SetSheetName;
    property NumberMask: TNumberMask read FNumberMask write FNumberMask;
    //---------events-----------//
    property OnBeforeStartFile: TBeforeStartFile read FOnBeforeStartFile write FOnBeforeStartFile;
    property OnAfterStartFile : TAfterStartFile  read FOnAfterStartFile  write FOnAfterStartFile;

  public
    destructor Destroy; override;
    constructor Create(AOwner: TComponent); override;
    function StartSheet: TOpenOffice_Calc;
    function AddNewSheet(const ASheetName: string; APosition: Integer): TOpenOffice_Calc;
    function PositionSheetByIndex(const ASheetIndex: Integer): TOpenOffice_Calc;
    function PositionSheetByName(const ASheetName: string):TOpenOffice_Calc;
    function SetFormula(ACellNumber: Integer; const aCollName: string; const aFormula: string): TOpenOffice_Calc;
    function SetValue(ACellNumber: Integer; const aCollName: string; aValue: variant; TypeValue: TTypeValue = ftString; AWrapped: boolean = false): TOpenOffice_Calc;
    function GetValue(ACellNumber: Integer; const aCollName: String) : TOpenOffice_Calc;
    function DataSetToSheet(const ACds : TClientDataSet): TOpenOffice_Calc;
    function SheetToDataSet(const ATabSheetName: String): TClientDataSet;
    
    procedure CallConversorPDFTOSheet;
    procedure ExeThread(AProc : TProc);

    property Value : string read FValue write FValue;
  end;

procedure Register;

implementation

uses
  Math,
  uOpenOffice.Helpers,
  uOpenOffice.Collors,
  uConvert.PDFToSheet;

procedure Register;
begin
  RegisterComponents('DinosOffice', [TOpenOffice_Calc]);
end;

procedure TOpenOffice_Calc.CallConversorPDFTOSheet;
begin
  TConvertPDFToSheet.CallConversor;
end;

procedure TOpenOffice_Calc.ValidateSheetName;
var
  lCID: LangID;
  lLanguages: array [0 .. 100] of char;
  lLanguage : string;
begin

  lCID := GetSystemDefaultLangID;

  if SheetName.Trim.IsEmpty then
  begin

    VerLanguageName(lCID, lLanguages, 100);
    lLanguage := String(lLanguages);

    if pos('Português', lLanguage) > 0 then
      SheetName := DefaultNewSheetNamePT
    else
      SheetName := DefaultNewSheetNameEn;
  end;
end;

function TOpenOffice_Calc.SetValue(ACellNumber: integer; const ACollName: string; AValue: variant; TypeValue: TTypeValue; AWrapped: boolean): TOpenOffice_Calc;
var
  lMap: string;
begin
  if ACellNumber = 0 then
    ACellNumber := 1;

  lMap := ACollName + ACellNumber.ToString;
  ObjCell := objSCalc.getCellRangeByName(lMap);

  if  Assigned(OnBeforeSetValue) then
    OnBeforeSetValue(Self);

  if TypeValue = ftString then
  begin
    ObjCell.IsTextWrapped := false;

    if AWrapped then
      ObjCell.IsTextWrapped := True;

    ObjCell.SetString(AValue);
  end
  else
  begin
    ObjCell.NumberFormat := Integer(NumberMask);
    ObjCell.SetValue(AValue);
  end;

  if  Assigned(OnAfterSetValue) then
    OnAfterSetValue(Self);

  Result := Self;    
end;

constructor TOpenOffice_Calc.Create(AOwner: TComponent);
begin
  inherited;
  Fields.setArrayFieldsSheet;
end;

procedure TOpenOffice_Calc.SetSheetName(const Value: string);
begin
  FSheetName := Value;
end;

function TOpenOffice_Calc.DataSetToSheet(const ACds: TClientDataSet): TOpenOffice_Calc;
var lIdx,lIdxFields : integer;
    lTypeVl : TTypeValue;
begin
  ACds.DisableControls;
  try
    //Create header
    for lIdx := 0 to pred(ACds.Fields.Count) do
      SetValue(0,FFields.ArrFields[lIdx],ACds.Fields[lIdx].DisplayName)
        .SetBold(True)
        .SetBorder([bAll], opBlack)
        .ChangeFont('Liberation Sans',11)
        .SetColor(opWhite, opSoftGray);

      ACds.First;
      while not ACds.Eof do
      begin
        for lIdxFields := 0 to pred(ACds.Fields.Count) do
        begin
          if (ACds.Fields[lIdxFields] is TCurrencyField) or
             (ACds.Fields[lIdxFields] is TIntegerField)  or
             (ACds.Fields[lIdxFields] is TFloatField)    or
             (ACds.Fields[lIdxFields] is TNumericField)  then
            lTypeVl := ftNumeric
           else
             lTypeVl := ftString;

          SetValue(ACds.RecNo +1, FFields.ArrFields[lIdxFields],ACds.Fields[lIdxFields].Value, lTypeVl)
          .SetBorder([bAll], opBlack);
        end;
        ACds.Next;
      end;
  finally
     ACds.EnableControls;
  end;

  Result := Self;
end;

destructor TOpenOffice_Calc.Destroy;
begin
  inherited;
end;

procedure TOpenOffice_Calc.ExeThread(AProc: TProc);
begin
  HungThread.ExecProc := AProc;
  HungThread.Start;
end;

function TOpenOffice_Calc.GetValue(ACellNumber: Integer; const ACollName: String) : TOpenOffice_Calc;
var
  lMap: string;
begin
  if Assigned(OnBeforeGetValue) then
    OnBeforeGetValue(self);


  lMap := ACollName + ACellNumber.ToString;
  ObjCell := objSCalc.getCellRangeByName(lMap);
  Value := VarToStr(ObjCell.String);

  Result := Self;

  if Assigned(OnAfterGetValue) then
    OnAfterGetValue(self);
end;

function TOpenOffice_Calc.PositionSheetByName(const ASheetName: string):TOpenOffice_Calc;
begin
  ObjSCalc := ObjDocument.Sheets.getByName(ASheetName);
  Result := Self;
end;

function TOpenOffice_Calc.PositionSheetByIndex(const ASheetIndex: integer) :TOpenOffice_Calc;
begin
  ObjSCalc := ObjDocument.Sheets.getByIndex(ASheetIndex);
  Result := Self;
end;

function TOpenOffice_Calc.AddNewSheet(const ASheetName: string; APosition: integer): TOpenOffice_Calc;
begin
  ObjDocument.Sheets.insertNewByName(ASheetName, APosition);
  ObjSCalc := ObjDocument.Sheets.getByName(ASheetName);

  Result := Self;
end;

function TOpenOffice_Calc.SetFormula(ACellNumber: Integer; const ACollName: string;
  const AFormula: string): TOpenOffice_Calc;
var
  lMap: string;
begin
  lMap := ACollName + ACellNumber.ToString;
  ObjCell := objSCalc.getCellByPosition(Fields.getIndex(ACollName), ACellNumber);
  //ObjCell.Formula := AFormula;
  ObjCell.FormulaLocal  := AFormula;
  
  Result := Self;
end;

function TOpenOffice_Calc.StartSheet : TOpenOffice_Calc;
begin
  if Assigned( FOnBeforeStartFile) then
    FOnBeforeStartFile(self);

  if URlFile.Trim.IsEmpty then
    URlFile := NewFile[integer(TpCalc)];

  ValidateSheetName;
  LoadDocument(SheetName);

  if ObjDocument.Sheets.hasByName(SheetName) then
    ObjSCalc := ObjDocument.Sheets.getByName(SheetName)
  else
  begin
    ObjSCalc := ObjDocument.createInstance('com.sun.star.sheet.Spreadsheet');
    ObjDocument.Sheets.insertByName(SheetName, ObjSCalc);
  end;

  if Assigned( FOnAfterStartFile) then
     FOnAfterStartFile(self);

  Result := Self;
end;

function TOpenOffice_Calc.SheetToDataSet(const ATabSheetName: String): TClientDataSet;
var lIdx, lIdxField : Integer;
begin
  Result := TClientDataSet.Create(nil);
  try
     positionSheetByName(ATabSheetName);
     for lIdx := 0 to CountCell -1 do
       Result.FieldDefs.Add(GetValue(1,Fields.getField(lIdx)).Value,TFieldType.ftString,100);
     Result.CreateDataSet;
     Result.DisableControls;
     Result.LogChanges := false;
     for lIdx := 2 to CountRow do
     begin
       Result.Append;
       for lIdxField := 0 to pred(Result.FieldCount) do
         Result.Fields[lIdxField] .AsString := GetValue(lIdx,Fields.getField(lIdxField)).Value;

       Result.Post;
      end;
  finally
    Result.EnableControls;
    CloseFile;
  end;
end;

{ TFieldsSheet }
function TFieldsSheet.GetField(AIndex: integer): string;
var lDifIdx : double;
    lLetter : String;
begin

  if (AIndex > High(ArrFields) ) and (ArrFields[AIndex].Trim.IsEmpty ) then
  begin
     lDifIdx := AIndex / 26;
     lDifIdx := round(lDifIdx - 1);

     SetLength(ArrFields,AIndex);
  end;

  if ArrFields[AIndex].Trim.IsEmpty then
  begin
     lLetter := ArrFields[trunc(lDifIdx)]; //First lLetter

     if lDifIdx = 0 then
       lDifIdx := 1;

     lDifIdx := lDifIdx * 26;
     lLetter := lLetter + ArrFields[trunc(AIndex  - lDifIdx)]; //Other lLetter of collumn

    ArrFields[AIndex] := lLetter;
  end;


  Result := ArrFields[AIndex];
end;

function TFieldsSheet.GetIndex(ANameField: String): integer;
var
  lIndex, lIdx: integer;
  lRep,
  lAux,
  lFirstIdx,
  lSecondIdx : integer;
begin
  Result := 0;
  lRep := 26;
  lAux := 0;
  lSecondIdx := 0;
  ANameField := ANameField.ToUpper;
  if ANameField.length <= 1  then
  begin
    for lIndex := 0 to High(arrFields) do
      if arrFields[lIndex] = ANameField then
      begin
        Result := lIndex;
        exit;
      end;
  end else
  begin
    for lIdx := 1 to ANameField.Length - 2 do
      lRep := (lRep * 26) + 26;

      //26*26 + 26

    for lIdx := 2 to ANameField.Length do
    begin
      for lIndex := 0 to High(arrFields) do
          if arrFields[lIndex] = ANameField[lIdx] then
            break;
    end;
    lFirstIdx  :=  getIndex(ANameField[1]) + 1;

    if  ANameField.Length = 3 then
    begin
      lSecondIdx := getIndex(ANameField[2]) + 27;
      lAux := 26;
    end;

    Result := ( (lFirstIdx * 26) + (lSecondIdx * 26) + lIndex) - lAux;
  end;
end;

procedure TFieldsSheet.SetArrayFieldsSheet;
begin
  SetLength(ArrFields, 26);

  ArrFields[0] := 'A';
  ArrFields[1] := 'B';
  ArrFields[2] := 'C';
  ArrFields[3] := 'D';
  ArrFields[4] := 'E';
  ArrFields[5] := 'F';
  ArrFields[6] := 'G';
  ArrFields[7] := 'H';
  ArrFields[8] := 'I';
  ArrFields[9] := 'J';
  ArrFields[10] := 'K';
  ArrFields[11] := 'L';
  ArrFields[12] := 'M';
  ArrFields[13] := 'N';
  ArrFields[14] := 'O';
  ArrFields[15] := 'P';
  ArrFields[16] := 'Q';
  ArrFields[17] := 'R';
  ArrFields[18] := 'S';
  ArrFields[19] := 'T';
  ArrFields[20] := 'U';
  ArrFields[21] := 'V';
  ArrFields[22] := 'W';
  ArrFields[23] := 'X';
  ArrFields[24] := 'Y';
  ArrFields[25] := 'Z';
end;

end.
