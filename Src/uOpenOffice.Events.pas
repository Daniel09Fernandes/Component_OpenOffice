{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOfficeEvents.pas }
{ Developer: Daniel Fernandes Rodrigures }
{ Email    : danielfernandesroddrigues@gmail.com }
{ this unit is a part of the Open Source. }
{ licensed under a MIT - see LICENSE.md}

{ ******************************************************* }

unit uOpenOffice.Events;

interface
 uses classes, uOpenOffice.SetPrinter;

 type TBeforeStartFile  = procedure (sender : TObject) of object;
 type TAfterStartFile   = procedure (sender : TObject) of object;
 type TBeforeCloseFile  = procedure (sender : TObject) of object;
 type TAfterCloseFile   = procedure (sender : TObject) of object;
 type TBeforePrint      = procedure (sender : TObject; var SetPrinter: TSetPrinter ) of object;
 type TAfterSetValue    = procedure (sender : TObject) of object;
 type TBeforeSetValue   = procedure (sender : TObject) of object;
 type TAfterGetValue    = procedure (sender : TObject) of object;
 type TBeforeGetValue   = procedure (sender : TObject) of object;
implementation

end.
