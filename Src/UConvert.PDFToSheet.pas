unit UConvert.PDFToSheet;

interface

uses
  SHDocVw, Vcl.Controls;

type
  TConvertPDFToSheet = class
  private
    Const
      FUrl = 'https://www.ilovepdf.com/pt/pdf_para_excel';
  public
    class procedure CallConversor;
  end;

implementation

uses shellApi;
{ TConvertPDFToSheet }

class procedure TConvertPDFToSheet.CallConversor;
begin
//Provisorio
  ShellExecute(0,'OPEN', FUrl, '', '', 0);
end;

end.
