unit MainModule;

interface

uses
  uniGUIMainModule, SysUtils, Classes, uOpenOffice, uOpenOffice.Calc;

type
  TUniMainModule = class(TUniGUIMainModule)
    OpenOffice_calc1: TOpenOffice_calc;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

function UniMainModule: TUniMainModule;

implementation

{$R *.dfm}

uses
  UniGUIVars, ServerModule, uniGUIApplication;

function UniMainModule: TUniMainModule;
begin
  Result := TUniMainModule(UniApplication.UniMainModule)
end;

initialization
  RegisterMainModuleClass(TUniMainModule);
end.
