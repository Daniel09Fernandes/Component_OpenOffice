
unit uOpenOffice.HungThread;

interface

uses
  ActiveX, System.Variants,System.SysUtils,
  Windows, System.SyncObjs, System.Classes;

type
  TOpenOfficeHungThread = class(TThread)
  private
    FProc: Tproc;
    procedure ExeProcedureExternal;
  public
   constructor Create;
   destructor Destroy; override;

   procedure Execute; override;

   property ExecProc: Tproc read FProc write FProc;

 end;

implementation

uses
  Vcl.Forms;

constructor TOpenOfficeHungThread.Create;
begin
  inherited Create(true);
end;

destructor TOpenOfficeHungThread.Destroy;
begin
  inherited;
end;

procedure TOpenOfficeHungThread.Execute;
begin
  inherited;
  Self.Queue(self.ExeProcedureExternal);
end;

procedure TOpenOfficeHungThread.ExeProcedureExternal;
begin
 if Assigned(FProc) then
   FProc();
end;

end.
