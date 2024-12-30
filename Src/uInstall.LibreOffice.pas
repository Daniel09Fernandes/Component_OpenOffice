unit uInstall.LibreOffice;

interface

uses  SysUtils, Variants, Classes, Vcl.OleCtrls, SHDocVw, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TInstallLibreOffice = class(TComponent)
  private
    const
      DonwloadClickJS = ' document.getElementsByClassName("dl_download_link")[0].click(); ';
      AURL = 'https://www.libreoffice.org/download/download-libreoffice/';
    var
      FURL_download :string;
      FWebBrowser : TWebBrowser;

    procedure Download;
    procedure WhenDocIsCompleted(ASender: TObject; const ADisp: IDispatch; const AURL: OleVariant);
  public
     constructor Create(AOwner: TComponent); override;
     destructor Destroy; override;

     procedure DownloadLibreOffice;
  end;


implementation

uses
  Vcl.Controls, MSHTML, Vcl.Forms;

{ TInstallLibreOffice }

constructor TInstallLibreOffice.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FURL_download := 'https://www.libreoffice.org/donate/dl/win-x86_64/versao/pt-BR/LibreOffice_versao_Win_x64.msi';

  if not assigned(FWebBrowser) then
    FWebBrowser := TWebBrowser.Create(self);
end;

destructor TInstallLibreOffice.Destroy;
begin
  FreeAndNil(FWebBrowser);
  inherited;
end;

procedure TInstallLibreOffice.Download;
begin
   FWebBrowser.Navigate(FURL_download);
   FWebBrowser.OnDocumentComplete := nil;
end;

procedure TInstallLibreOffice.DownloadLibreOffice;
begin
  TWinControl(FWebBrowser).Name   := 'WebBrowser';
  FWebBrowser.Silent := true;  //don't show JS errors
  FWebBrowser.Visible:= false;  //visible...by default true
  FWebBrowser.HandleNeeded;

  FWebBrowser.Navigate(AURL);
  FWebBrowser.RegisterAsBrowser:= True;
  FWebBrowser.OnDocumentComplete :=  WhenDocIsCompleted;
  FWebBrowser.Top    := 0;
  FWebBrowser.Left   := 0;
  FWebBrowser.Height := 0;
  FWebBrowser.Width  := 0;
end;

procedure TInstallLibreOffice.WhenDocIsCompleted(ASender: TObject; const ADisp: IDispatch; const AURL: OleVariant);
var
 lDoc: IHTMLDocument2;
 lBody : string;
 lVersao : string;
 lUrlDownload: string;
begin
  lDoc := FWebBrowser.Document as IHTMLDocument2;
  lBody :=   lDoc.Body.innerText;
  lVersao := Copy(lBody,pos('is available for the following operating systems/architectures',lBody)-6,6).Trim;
  lUrlDownload := FURL_download;
  lUrlDownload := lUrlDownload.Replace('versao',lVersao);
  FURL_download :=  lUrlDownload;
  Download;
end;

end.
