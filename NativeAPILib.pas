unit NativeAPILib;

interface

type
  TGetClassNamesFunc = function(): PWideChar; cdecl;
  TGetClassObjectFunc = function(const clsName: PWideChar; pIntf: Pointer): Integer; cdecl;
  TDestroyObjectFunc = function(pIntf: Pointer): Integer; cdecl;

implementation

end.
