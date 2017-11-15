unit AddInDefBase;

interface

uses
  V8Types;

type
  { IAddInDefBase }
  TAddErrorFunc = function(wcode: UInt16; const source: PWideChar; const descr: PWideChar; scode: LongInt): Boolean of object; stdcall;
  TReadFunc = function(wszPropName: PWideChar; pVal: P1CVariant; pErrCode: Pointer; errDescriptor: PPWideChar): Boolean of object; stdcall;
  TWriteFunc = function(wszPropName: PWideChar; pVar: P1CVariant): Boolean of object; stdcall;
  TRegisterProfileAsFunc = function(wszProfileName: PWideChar): Boolean of object; stdcall;
  TSetEventBufferDepthFunc = function(lDepth: LongInt): Boolean of object; stdcall;
  TGetEventBufferDepthFunc = function(): Boolean of object; stdcall;
  TExternalEventFunc = function(wszSource: PWideChar; wszMessage: PWideChar; wszData: PWideChar): Boolean of object; stdcall;
  TCleanEventBufferFunc = procedure() of object; stdcall;
  TSetStatusLineFunc = function(wszStatusLine: PWideChar): Boolean of object; stdcall;
  TResetStatusLineFunc = procedure() of object; stdcall;

  P1CAddInDefBase = ^T1CAddInDefBase;
  P1CAddInDefBaseVTable = ^T1CAddInDefBaseVTable;
  T1CAddInDefBase = packed record
    __vfptr: P1CAddInDefBaseVTable;
  end;

  T1CAddInDefBaseVTable = packed record
    _Destructor: procedure(This: P1CAddInDefBase); cdecl;
    AddError: TAddErrorFunc;
    Read: TReadFunc;
    Write: TWriteFunc;
    RegisterProfileAs: TRegisterProfileAsFunc;
    SetEventBufferDepth: TSetEventBufferDepthFunc;
    GetEventBufferDepth: TGetEventBufferDepthFunc;
    ExternalEvent: TExternalEventFunc;
    CleanEventBuffer: TCleanEventBufferFunc;
    SetStatusLine: TSetStatusLineFunc;
    ResetStatusLine: TResetStatusLineFunc;
  end;

  TAddInDefBase = class
  private
    FAddInDefBaseRec: T1CAddInDefBase;
  public
    constructor Create();
    destructor Destroy(); override;

    property AddInDefBaseRec: T1CAddInDefBase read FAddInDefBaseRec;
  end;

implementation

constructor TAddInDefBase.Create;
begin
  inherited;

  GetMem(FAddInDefBaseRec.__vfptr, SizeOf(T1CAddInDefBaseVTable));
  FAddInDefBaseRec.__vfptr._Destructor := nil;
  FAddInDefBaseRec.__vfptr.AddError := nil;
  FAddInDefBaseRec.__vfptr.Read := nil;
  FAddInDefBaseRec.__vfptr.Write := nil;
  FAddInDefBaseRec.__vfptr.RegisterProfileAs := nil;
  FAddInDefBaseRec.__vfptr.SetEventBufferDepth := nil;
  FAddInDefBaseRec.__vfptr.GetEventBufferDepth := nil;
  FAddInDefBaseRec.__vfptr.ExternalEvent := nil;
  FAddInDefBaseRec.__vfptr.CleanEventBuffer := nil;
  FAddInDefBaseRec.__vfptr.SetStatusLine := nil;
  FAddInDefBaseRec.__vfptr.ResetStatusLine := nil;
end;

destructor TAddInDefBase.Destroy;
begin
  FreeMem(FAddInDefBaseRec.__vfptr);
end;

end.
