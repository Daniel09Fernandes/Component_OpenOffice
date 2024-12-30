{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOfficeSetPrinter.pas }
{ Developer: Daniel Fernandes Rodrigures }
{ Email    : danielfernandesroddrigues@gmail.com }
{ this unit is a part of the Open Source. }
{ licensed under a MIT - see LICENSE.md}

{ ******************************************************* }

unit uOpenOffice.SetPrinter;

interface
uses
  classes;

type TSetPrinter = class(TCollectionItem)
   private
    FPaperSize_Height: integer;
    FPages: string;
    FPaperSize_Width: integer;
    FPrinterName: string;
    procedure SetPages(const Value: string);
    procedure SetPaperSize_Height(const Value: integer);
    procedure SetPaperSize_Width(const Value: integer);
    procedure SetPrinterName(const Value: string);

   public
    constructor Create(AOwner: TCollection); override;
    property PaperSize_Width: integer read FPaperSize_Width write SetPaperSize_Width; //20000   ' corresponds to 20 cm
    property PaperSize_Height: integer read FPaperSize_Height write SetPaperSize_Height;
    property PrinterName: string read FPrinterName write SetPrinterName;
    property Pages : string read FPages write SetPages; //1-3; 7; 9;
 end;

implementation

{ TSetPrinter }

constructor TSetPrinter.Create(AOwner: TCollection);
begin
  inherited;
  FPaperSize_Height := 20000;
  FPaperSize_Width := 20000;
  FPages :=  '1-10;';
  FPrinterName := 'Microsoft Print to PDF';
end;

procedure TSetPrinter.SetPages(const Value: string);
begin
  FPages := Value;
end;

procedure TSetPrinter.SetPaperSize_Height(const Value: integer);
begin
  FPaperSize_Height := Value;
end;

procedure TSetPrinter.SetPaperSize_Width(const Value: integer);
begin
  FPaperSize_Width := Value;
end;

procedure TSetPrinter.SetPrinterName(const Value: string);
begin
  FPrinterName := Value;
end;

end.
