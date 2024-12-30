{MIT License

Copyright (c) 2022 Daniel Fernandes

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.}
{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOffice.pas }
{ Developer: Daniel Fernandes Rodrigures }
{ Email    : danielfernandesroddrigues@gmail.com }
{ this unit is a part of the Open Source. }
{ licensed under a MIT - see LICENSE.md}

{ ******************************************************* }

unit uOpenOffice;

interface

uses ActiveX, System.Classes, Vcl.Dialogs, System.Variants, Windows, System.UITypes,
  uOpenOffice.SetPrinter,
  uOpenOffice.Events,
  uInstall.LibreOffice,
  uOpenOffice.HungThread;

Type
  TTypeOffice = (TpCalc, TpWriter);
  TTypeLanguage = (TpCalcPTbr, TpCalcEn, TpWriterPTbr, TpWriterEn);

Type
  TOpenOffice = class(TComponent)
  private
    FURlFile: string;
    FSetPrinter: TSetPrinter;
    FOpenOfficeHungThread : TOpenOfficeHungThread;
    FOnBeforePrint: TBeforePrint;
    FOnBeforeCloseFile: TBeforeCloseFile;
    FOnAfterCloseFile: TAfterCloseFile;
    FOnAfterGetValue: TAfterGetValue;
    FOnBeforeGetValue: TBeforeGetValue;
    FOnAfterSetValue: TAfterSetValue;
    FOnBeforeSetValue: TBeforeSetValue;
    InstallLibreOffice: TInstallLibreOffice;
    FDocVisible: Boolean;

    procedure SetURlFile(const Value: string);
    procedure Inicialization;
    procedure SetParamsInicialization;
  protected
    { Protected declarations }
    ObjCoreReflection,
    ObjDesktop,
    ObjServiceManager,
    ObjDocument,
    OValMacro,
    ObjSCalc,
    ObjWriter,
    ObjDispatcher,
    ObjCell,
    ObjCharts: OleVariant;
    OInicializationProperties : array [0 .. 1] of Variant;
    NewFile: array [0 .. 1] of string;

    function ConvertFilePathToUrlFile(AFilePath: string): string;
    procedure LoadDocument(AFileName: string = '');

    Property SetPrinter: TSetPrinter read FSetPrinter write FSetPrinter;
    property HungThread : TOpenOfficeHungThread read FOpenOfficeHungThread write FOpenOfficeHungThread;
  published
    property URlFile: string read FURlFile write SetURlFile;
    property OnBeforePrint: TBeforePrint read FOnBeforePrint
      write FOnBeforePrint;
    property OnBeforeCloseFile: TBeforeCloseFile read FOnBeforeCloseFile
      write FOnBeforeCloseFile;
    property OnAfterCloseFile: TAfterCloseFile read FOnAfterCloseFile
      write FOnAfterCloseFile;
    property OnBeforeGetValue: TBeforeGetValue read FOnBeforeGetValue
      write FOnBeforeGetValue;
    property OnAfterGetValue: TAfterGetValue read FOnAfterGetValue
      write FOnAfterGetValue;
    property OnBeforeSetValue: TBeforeSetValue read FOnBeforeSetValue
      write FOnBeforeSetValue;
    property OnAfterSetValue: TAfterSetValue read FOnAfterSetValue
      write FOnAfterSetValue;
    property DocVisible : boolean read FDocVisible write FDocVisible;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure Print;
    procedure CloseFile;
    procedure SaveFile(AFileName: String);
  end;

implementation

uses
  System.SysUtils, System.Win.ComObj;

{ TOpenOffice }

procedure TOpenOffice.CloseFile;
begin
  if Assigned(FOnBeforeCloseFile) then
    FOnBeforeCloseFile(Self);

  ObjDocument.close(True);

  if Assigned(FOnAfterCloseFile) then
    FOnAfterCloseFile(Self);
end;

function TOpenOffice.ConvertFilePathToUrlFile(AFilePath: string): string;
begin
  if (pos('FILE:///', UpperCase(AFilePath)) <= 0) then
  begin
    AFilePath := StringReplace(AFilePath, '\', '/', [rfReplaceAll]);
    AFilePath := 'file:///' + AFilePath;
  end;
  Result := AFilePath;
end;

procedure TOpenOffice.SetURlFile(const Value: string);
begin
  FURlFile := Value;

  if FURlFile.Trim.IsEmpty or (FURlFile = NewFile[Integer(TpCalc)]) or
    (FURlFile = NewFile[Integer(TpWriter)]) then
    exit;

  FURlFile := ConvertFilePathToUrlFile(FURlFile);
end;

constructor TOpenOffice.Create(AOwner: TComponent);
procedure InitializeCOM;
begin
  try
    if TThread.CurrentThread.ThreadID = MainThreadID then
    begin
      OleCheck(CoInitialize(nil));
    end
    else
    begin
      // Para APIs multithread, Unigui, Intraweb, inicialize com COINIT_MULTITHREADED
      try
        OleCheck(CoInitializeEx(nil, COINIT_MULTITHREADED));
      except
        OleCheck(CoInitialize(nil));
      end;
    end;
  except
    on E: Exception do
      raise Exception.Create('Erro ao inicializar COM: ' + E.Message);
  end;
end;
begin
  inherited;
  InitializeCOM;
  FOpenOfficeHungThread := TOpenOfficeHungThread.Create;
  Inicialization;
  FSetPrinter := TSetPrinter.Create(nil);
  NewFile[Integer(TpCalc)] := 'private:factory/scalc';
  NewFile[Integer(TpWriter)] := 'private:factory/swriter';
end;

destructor TOpenOffice.Destroy;
begin
  inherited;
  FSetPrinter.Free;
  ObjCoreReflection := Unassigned;
  ObjDesktop := Unassigned;
  ObjServiceManager := Unassigned;
  ObjDocument := Unassigned;
  ObjSCalc := Unassigned;
  ObjCell := Unassigned;

  freeAndNil(FOpenOfficeHungThread);
  if assigned(InstallLibreOffice) then
    freeAndNil(InstallLibreOffice);
end;

procedure TOpenOffice.Inicialization;
begin
  try
    // Libre office
    ObjServiceManager := CreateOleObject('com.sun.star.ServiceManager');
    ObjCoreReflection := ObjServiceManager.createInstance
      ('com.sun.star.reflection.CoreReflection');
    ObjDesktop := ObjServiceManager.createInstance('com.sun.star.frame.Desktop');
  except
    if messageDlg('Erro(pt-Br):  Instale o LibreOffice para usar o sistema' +
      #13 + #13 + 'Error(En)  :  install  the LibreOffice to use the system' +
      #13#13 + 'Dowload in: https://www.libreoffice.org/download/download-libreoffice/'
      + #13#13 + '(pt-Br) Deseja instalar a versão mais recente do LibreOffice?'
      + #13 + '(En) Do you want to install the latest version of LibreOffice?',

      TMsgDlgType.mtWarning, [TMsgDlgBtn.mbYes, TMsgDlgBtn.mbNo], 0) = mrYes
    then
    begin
      InstallLibreOffice := TInstallLibreOffice.Create(nil);
      InstallLibreOffice.DownloadLibreOffice;
    end;
  end;
end;

procedure TOpenOffice.LoadDocument(AFileName: string = '');
var lIdx: Integer;
begin
  CoInitialize(nil);
  if AFileName = '' then
    AFileName := '_blank';

  for lIdx := 0 to High(OInicializationProperties) do
    VarClear(OInicializationProperties[lIdx]);

  if not DocVisible then
    SetParamsInicialization;

  ObjDocument := ObjDesktop.loadComponentFromURL(URlFile, AFileName, 0,VarArrayOf(OInicializationProperties));
end;

procedure TOpenOffice.SetParamsInicialization;
begin
    OInicializationProperties[0] := ObjServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
    OInicializationProperties[0].Name := 'Hidden';
    OInicializationProperties[0].Value := true;

    OValMacro :=  ObjServiceManager.createInstance('com.sun.star.document.MacroExecMode.ALWAYS_EXECUTE_NO_WARN');

    OInicializationProperties[1] := ObjServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
    OInicializationProperties[1].Name := 'MacroExecutionMode';
    OInicializationProperties[1].Value := OValMacro;
end;

procedure TOpenOffice.Print;
var
  lPaperSize: Variant;
  lPrinterProperties: array [0 .. 3] of Variant;
begin

  if Assigned(OnBeforePrint) then
  begin
    OnBeforePrint(Self, FSetPrinter);

    lPaperSize := ObjServiceManager.Bridge_GetStruct('com.sun.star.awt.Size');

    lPaperSize.Width := FSetPrinter.PaperSize_Width;
    lPaperSize.Height := FSetPrinter.PaperSize_Height;

    lPrinterProperties[0] := ObjServiceManager.Bridge_GetStruct
      ('com.sun.star.beans.PropertyValue');
    lPrinterProperties[0].Name := 'Name';
    lPrinterProperties[0].Value := FSetPrinter.PrinterName;

    lPrinterProperties[1] := ObjServiceManager.Bridge_GetStruct
      ('com.sun.star.beans.PropertyValue');
    lPrinterProperties[1].Name := 'PaperSize';
    lPrinterProperties[1].Value := lPaperSize;

    lPrinterProperties[2] := ObjServiceManager.Bridge_GetStruct
      ('com.sun.star.beans.PropertyValue');
    lPrinterProperties[2].Name := 'Pages';
    lPrinterProperties[2].Value := FSetPrinter.Pages;

    ObjDocument.Printer := VarArrayOf(lPrinterProperties);

    ObjDocument.print(VarArrayOf(lPrinterProperties));
  end
  else
    ObjDocument.print(VarArrayOf([]));
end;

procedure TOpenOffice.SaveFile(AFileName: String);
var
  lSaveProperty : array [0..1] of Variant;
begin
  AFileName := ConvertFilePathToUrlFile(AFileName);

  if AFileName.Contains('.xlsx') then
  begin
    //Codigo fornecido por @adolfomayer - e adptado por @dinosdev 29/11/2024
    lSaveProperty[0] := ObjServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
    lSaveProperty[0].Name := 'FilterName';
    lSaveProperty[0].Value := 'Calc MS Excel 2007 XML'; //for XLSX

    lSaveProperty[1] := ObjServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
    lSaveProperty[1].Name := 'Overwrite';
    lSaveProperty[1].Value := True;
    ObjDocument.storeAsURL(AFileName, VarArrayOf(lSaveProperty))
  end
  else
    ObjDocument.storeAsURL(AFileName, VarArrayOf([]));
end;

end.
