library ATOLWrapper;

{$R *.res}
{$R 'WrapperObject.res' 'WrapperObject.rc'}

uses
  ComServ,
  WrapperObject in 'WrapperObject.pas';

{$E dll}

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

begin
end.
