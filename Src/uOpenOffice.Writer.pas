{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOffice_writer.pas }
{ Developer: Daniel Fernandes Rodrigures }
{ Email    : danielfernandesroddrigues@gmail.com }
{ this unit is a part of the Open Source. }
{ licensed under a MIT - see LICENSE.md}

{ ******************************************************* }
unit uOpenOffice.Writer;

interface

uses
  System.Classes, DB, ActiveX, uOpenOffice,
  dbWeb, ComObj, XMLDoc, XMLIntf, Vcl.Dialogs, System.Variants, Windows,
  uOpenOffice.Events,
  uOpenOffice.SetPrinter,
  uOpenOffice.Collors;

type
  // TTypeValue = (ftString, ftNumeric);
  TOpenOffice_writer = class(TOpenOffice)
  private
    objTextCursor, oText, oCursor: variant;
    FDocName,
    FValueText: string;
    FBoldActive,
    changeForDispatcher: boolean;
    procedure SetDocName(const Value: string);
  public
  var
    Value: string;
    propsText: array [0 .. 4] of variant;
    destructor Destroy; override;
    constructor Create(AOwner: TComponent); override;
    procedure startDoc;
    function gotoEndOfSentence : TOpenOffice_writer;
    function setValue(const aText: string): TOpenOffice_writer;
    function getValue: string;
    property BoldActive : boolean read FBoldActive write FBoldActive;
    property Cursor : variant read oCursor;
    property ValueText : string read FValueText;
  published
    property ServicesManager: OleVariant read objServiceManager;
    property DocName: string read FDocName write SetDocName;
  end;

  procedure Register;

implementation

{ TOpenOffice_writer }
uses
  System.SysUtils, System.Math;

procedure Register;
begin
  RegisterComponents('DinosOffice', [TOpenOffice_writer]);
end;

constructor TOpenOffice_writer.Create(AOwner: TComponent);
begin
  inherited;
  FBoldActive := false;
end;

destructor TOpenOffice_writer.Destroy;
var
  i: integer;
begin
  for i := 0 to High(propsText) do
    propsText[i] := Unassigned;

  inherited;
end;

function TOpenOffice_writer.getValue: string;
begin
  if assigned(onBeforeGetValue) then
    onBeforeGetValue(self);

  Result := VarToStr(oText.String);

  if assigned(onAfterGetValue) then
    onAfterGetValue(self);
end;


function TOpenOffice_writer.gotoEndOfSentence: TOpenOffice_writer;
begin
  objTextCursor.jumpToEndOfPage;
  oText := objTextCursor.Text;
  oCursor := oText.CreateTextCursor;
  oCursor.gotoEnd(true);
  Result := self;
end;

function TOpenOffice_writer.setValue(const aText: string): TOpenOffice_writer;
begin

  if assigned(onBeforeSetValue) then
    onBeforeSetValue(self);


  oText := objTextCursor.Text;
  oText.InsertString(objTextCursor, '', true);

  oCursor := oText.CreateTextCursor;

  changeForDispatcher := true;

  oText.InsertString(objTextCursor, aText, true);
  oText.InsertControlCharacter(oCursor, 10, true);

  FValueText  := aText;
  Result      := self;

  if assigned(onAfterSetValue) then
    onAfterSetValue(self);
end;

procedure TOpenOffice_writer.startDoc;
begin
  if URlFile.Trim.IsEmpty then
    URlFile := NewFile[integer(TpWriter)];

  LoadDocument(DocName); // cria a instancia do objDocument
  objTextCursor := objDocument.getCurrentController.getViewCursor;

  propsText[0] := objServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[1] := objServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[2] := objServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[3] := objServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');
  propsText[4] := objServiceManager.Bridge_GetStruct
    ('com.sun.star.beans.PropertyValue');

  objWriter := objDocument.getCurrentController.getFrame;
  objDispatcher := objServiceManager.createInstance
    ('com.sun.star.frame.DispatchHelper');
end;

procedure TOpenOffice_writer.SetDocName(const Value: string);
begin
  FDocName := Value;
end;

end.
