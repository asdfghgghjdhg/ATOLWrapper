unit ComponentBase;

interface

uses
  V8Types;

type
  { IInitDoneBase }
  PInitDoneBase = ^TInitDoneBase;
  PInitDoneBaseVTable = ^TInitDoneBaseVTable;
  TInitDoneBase = packed record
    __vfptr: PInitDoneBaseVTable;
  end;

  TInitFunc = function(This: PInitDoneBase; disp: Pointer): longbool; stdcall;
  TSetMemManagerFunc = function(This: PInitDoneBase; mem: Pointer): longbool; stdcall;
  TGetInfoFunc = function(This: PInitDoneBase): LongInt; stdcall;
  TDoneFunc = procedure(This: PInitDoneBase); stdcall;

  TInitDoneBaseVTable = packed record
    _Destructor: procedure(This: PInitDoneBase); cdecl;
    Init: TInitFunc;
    setMemManager: TSetMemManagerFunc;
    GetInfo: TGetInfoFunc;
    Done: TDoneFunc;
  end;

  { ILanguageExtenderBase }
  PLanguageExtenderBase = ^TLanguageExtenderBase;
  PLanguageExtenderBaseVTable = ^TLanguageExtenderBaseVTable;
  TLanguageExtenderBase = packed record
    __vfptr: PLanguageExtenderBaseVTable;
  end;

  TRegisterExtensionAsFunc = function(This: PLanguageExtenderBase; var wsExtensionName: PWideChar): longbool; stdcall;
  TGetNPropsFunc = function(This: PLanguageExtenderBase): LongInt; stdcall;
  TFindPropFunc = function(This: PLanguageExtenderBase; const wsPropName: PWideChar): LongInt; stdcall;
  TGetPropNameFunc = function(This: PLanguageExtenderBase; lPropNum: LongInt; lPropAlias: LongInt): PWideChar; stdcall;
  TGetPropValFunc = function(This: PLanguageExtenderBase; const lPropNum: LongInt; pvarPropVal: P1CVariant): longbool; stdcall;
  TSetPropValFunc = function(This: PLanguageExtenderBase; const lPropNum: LongInt; varPropVal: P1CVariant): longbool; stdcall;
  TIsPropReadableFunc = function(This: PLanguageExtenderBase; const lPropNum: LongInt): longbool; stdcall;
  TIsPropWritableFunc = function(This: PLanguageExtenderBase; const lPropNum: LongInt): longbool; stdcall;
  TGetNMethodsFunc = function(This: PLanguageExtenderBase): LongInt; stdcall;
  TFindMethodFunc = function(This: PLanguageExtenderBase; const wsMethodName: PWideChar): LongInt; stdcall;
  TGetMethodNameFunc = function(This: PLanguageExtenderBase; const lMethodNum: LongInt; const lMethodAlias: LongInt): PWideChar; stdcall;
  TGetNParamsFunc = function(This: PLanguageExtenderBase; const lMethodNum: LongInt): LongInt; stdcall;
  TGetParamDefValueFunc = function(This: PLanguageExtenderBase; const lMethodNum: LongInt; const lParamNum: LongInt; pvarParamDefValue: P1CVariant): longbool; stdcall;
  THasRetValFunc = function(This: PLanguageExtenderBase; const lMethodNum: LongInt): longbool; stdcall;
  TCallAsProcFunc = function(This: PLanguageExtenderBase; const lMethodNum: LongInt; paParams: P1CVariant; const lSizeArray: LongInt): longbool; stdcall;
  TCallAsFuncFunc = function(This: PLanguageExtenderBase; const lMethodNum: LongInt; pvarRetValue: P1CVariant; paParams: P1CVariant; const lSizeArray: LongInt): longbool; stdcall;

  TLanguageExtenderBaseVTable = packed record
    _Destructor: procedure(This: PLanguageExtenderBase); cdecl;
    RegisterExtensionAs: TRegisterExtensionAsFunc;
    GetNProps: TGetNPropsFunc;
    FindProp: TFindPropFunc;
    GetPropName: TGetPropNameFunc;
    GetPropVal: TGetPropValFunc;
    SetPropVal: TSetPropValFunc;
    IsPropReadable: TIsPropReadableFunc;
    IsPropWritable: TIsPropWritableFunc;
    GetNMethods: TGetNMethodsFunc;
    FindMethod: TFindMethodFunc;
    GetMethodName: TGetMethodNameFunc;
    GetNParams: TGetNParamsFunc;
    GetParamDefValue: TGetParamDefValueFunc;
    HasRetVal: THasRetValFunc;
    CallAsProc: TCallAsProcFunc;
    CallAsFunc: TCallAsFuncFunc;
  end;

  { ILocaleBase }
  PLocaleBase = ^TLocaleBase;
  PLocaleBaseVTable = ^TLocaleBaseVTable;
  TLocaleBase = packed record
    __vfptr: PLocaleBaseVTable;
  end;

  TSetLocaleFunc = procedure(This: PLocaleBase; const loc: PWideChar); stdcall;

  TLocaleBaseVTable = packed record
    _Destructor: procedure(This: PLocaleBase); cdecl;
    SetLocale: TSetLocaleFunc;
  end;

  { IComponentBase }
  PComponentBase = ^TComponentBase;
  PComponentBaseVTable = ^TComponentBaseVTable;
  TComponentBase = packed record
    InitDoneBase: TInitDoneBase;
    LanguageExtenderBase: TLanguageExtenderBase;
    LocaleBase: TLocaleBase;
    __vfptr: PComponentBaseVTable;
  end;
  TComponentBaseVTable = packed record
    _Destructor: procedure(This: PComponentBase); cdecl;
  end;

  TNativeAPIClass = class
  private
    FComponentBaseRec: PComponentBase;
  public
    constructor Create();
    destructor Destroy(); override;

    function Init(disp: Pointer): Boolean;
    function SetMemManager(mem: Pointer): Boolean;
    function GetInfo(): LongInt;
    procedure Done();

    function RegisterExtensionAs(var wsExtensionName: PWideChar): Boolean;
    function GetNProps(): LongInt;
    function FindProp(const wsPropName: PWideChar): LongInt;
    function GetPropName(lPropNum: LongInt; lPropAlias: LongInt): PWideChar;
    function GetPropVal(const lPropNum: LongInt; pvarPropVal: P1CVariant): Boolean;
    function SetPropVal(const lPropNum: LongInt; varPropVal: P1CVariant): Boolean;
    function IsPropReadable(const lPropNum: LongInt): Boolean;
    function IsPropWritable(const lPropNum: LongInt): Boolean;
    function GetNMethods(): LongInt;
    function FindMethod(const wsMethodName: PWideChar): LongInt;
    function GetMethodName(const lMethodNum: LongInt; const lMethodAlias: LongInt): PWideChar;
    function GetNParams(const lMethodNum: LongInt): LongInt;
    function GetParamDefValue(const lMethodNum: LongInt; const lParamNum: LongInt; pvarParamDefValue: P1CVariant): Boolean;
    function HasRetVal(const lMethodNum: LongInt): Boolean;
    function CallAsProc(const lMethodNum: LongInt; paParams: P1CVariant; const lSizeArray: LongInt): Boolean;
    function CallAsFunc(const lMethodNum: LongInt; pvarRetValue: P1CVariant; paParams: P1CVariant; const lSizeArray: LongInt): Boolean;

    procedure SetLocale(const loc: PWideChar);

    property ComponentBaseRec: PComponentBase read FComponentBaseRec write FComponentBaseRec;
  end;

implementation

constructor TNativeAPIClass.Create;
begin
  inherited;

  FComponentBaseRec := nil;
end;

destructor TNativeAPIClass.Destroy;
begin
  FComponentBaseRec := nil;

  inherited Destroy;
end;

function TNativeAPIClass.Init(disp: Pointer): Boolean;
begin
  Result := false;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.InitDoneBase.__vfptr.Init) then Exit;

  try
    Result := FComponentBaseRec.InitDoneBase.__vfptr.Init(@FComponentBaseRec.InitDoneBase, disp);
  finally
  end;
end;

function TNativeAPIClass.SetMemManager(mem: Pointer): Boolean;
begin
  Result := false;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.InitDoneBase.__vfptr.setMemManager) then Exit;

  try
    Result := FComponentBaseRec.InitDoneBase.__vfptr.setMemManager(@FComponentBaseRec.InitDoneBase, mem);
  finally
  end;
end;

function TNativeAPIClass.GetInfo(): LongInt;
begin
  Result := 0;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.InitDoneBase.__vfptr.GetInfo) then Exit;

  try
    Result := FComponentBaseRec.InitDoneBase.__vfptr.GetInfo(@FComponentBaseRec.InitDoneBase);
  finally
  end;
end;

procedure TNativeAPIClass.Done;
begin
  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.InitDoneBase.__vfptr.Done) then Exit;

  try
    FComponentBaseRec.InitDoneBase.__vfptr.Done(@FComponentBaseRec.InitDoneBase);
  finally
  end;
end;

function TNativeAPIClass.RegisterExtensionAs(var wsExtensionName: PWideChar): Boolean;
begin
  Result := false;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.RegisterExtensionAs) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.RegisterExtensionAs(@FComponentBaseRec.LanguageExtenderBase, wsExtensionName);
  finally
  end;
end;

function TNativeAPIClass.GetNProps(): LongInt;
begin
  Result := 0;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.GetNProps) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.GetNProps(@FComponentBaseRec.LanguageExtenderBase);
  finally
  end;
end;

function TNativeAPIClass.FindProp(const wsPropName: PWideChar): LongInt;
begin
  Result := -1;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.FindProp) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.FindProp(@FComponentBaseRec.LanguageExtenderBase, wsPropName);
  finally
  end;
end;

function TNativeAPIClass.GetPropName(lPropNum: LongInt; lPropAlias: LongInt): PWideChar;
begin
  Result := nil;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.GetPropName) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.GetPropName(@FComponentBaseRec.LanguageExtenderBase, lPropNum, lPropAlias);
  finally
  end;
end;

function TNativeAPIClass.GetPropVal(const lPropNum: LongInt; pvarPropVal: P1CVariant): Boolean;
begin
  Result := false;
  pvarPropVal.vt := VTYPE_EMPTY;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.GetPropVal) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.GetPropVal(@FComponentBaseRec.LanguageExtenderBase, lPropNum, pvarPropVal);
  finally
  end;
end;

function TNativeAPIClass.SetPropVal(const lPropNum: LongInt; varPropVal: P1CVariant): Boolean;
begin
  Result := false;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.SetPropVal) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.SetPropVal(@FComponentBaseRec.LanguageExtenderBase, lPropNum, varPropVal);
  finally
  end;
end;

function TNativeAPIClass.IsPropReadable(const lPropNum: LongInt): Boolean;
begin
  Result := false;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.IsPropReadable) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.IsPropReadable(@FComponentBaseRec.LanguageExtenderBase, lPropNum);
  finally
  end;
end;

function TNativeAPIClass.IsPropWritable(const lPropNum: LongInt): Boolean;
begin
  Result := false;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.IsPropWritable) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.IsPropWritable(@FComponentBaseRec.LanguageExtenderBase, lPropNum);
  finally
  end;
end;

function TNativeAPIClass.GetNMethods(): LongInt;
begin
  Result := 0;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.GetNMethods) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.GetNMethods(@FComponentBaseRec.LanguageExtenderBase);
  finally
  end;
end;

function TNativeAPIClass.FindMethod(const wsMethodName: PWideChar): LongInt;
begin
  Result := 0;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.FindMethod) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.FindMethod(@FComponentBaseRec.LanguageExtenderBase, wsMethodName);
  finally
  end;
end;

function TNativeAPIClass.GetMethodName(const lMethodNum: LongInt; const lMethodAlias: LongInt): PWideChar;
begin
  Result := nil;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.GetMethodName) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.GetMethodName(@FComponentBaseRec.LanguageExtenderBase, lMethodNum, lMethodAlias);
  finally
  end;
end;

function TNativeAPIClass.GetNParams(const lMethodNum: LongInt): LongInt;
begin
  Result := 0;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.GetNParams) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.GetNParams(@FComponentBaseRec.LanguageExtenderBase, lMethodNum);
  finally
  end;
end;

function TNativeAPIClass.GetParamDefValue(const lMethodNum: LongInt; const lParamNum: LongInt; pvarParamDefValue: P1CVariant): Boolean;
begin
  Result := false;
  pvarParamDefValue.vt := VTYPE_EMPTY;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.GetParamDefValue) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.GetParamDefValue(@FComponentBaseRec.LanguageExtenderBase, lMethodNum, lParamNum, pvarParamDefValue);
  finally
  end;
end;

function TNativeAPIClass.HasRetVal(const lMethodNum: LongInt): Boolean;
begin
  Result := false;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.HasRetVal) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.HasRetVal(@FComponentBaseRec.LanguageExtenderBase, lMethodNum);
  finally
  end;
end;

function TNativeAPIClass.CallAsProc(const lMethodNum: LongInt; paParams: P1CVariant; const lSizeArray: LongInt): Boolean;
begin
  Result := false;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.CallAsProc) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.CallAsProc(@FComponentBaseRec.LanguageExtenderBase, lMethodNum, paParams, lSizeArray);
  finally
  end;
end;

function TNativeAPIClass.CallAsFunc(const lMethodNum: LongInt; pvarRetValue: P1CVariant; paParams: P1CVariant; const lSizeArray: LongInt): Boolean;
begin
  Result := false;
  pvarRetValue.vt := VTYPE_EMPTY;

  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LanguageExtenderBase.__vfptr.CallAsFunc) then Exit;

  try
    Result := FComponentBaseRec.LanguageExtenderBase.__vfptr.CallAsFunc(@FComponentBaseRec.LanguageExtenderBase, lMethodNum, pvarRetValue, paParams, lSizeArray);
  finally
  end;
end;

procedure TNativeAPIClass.SetLocale(const loc: PWideChar);
begin
  if not Assigned(FComponentBaseRec) then Exit;
  if not Assigned(FComponentBaseRec.LocaleBase.__vfptr.SetLocale) then Exit;

  try
    FComponentBaseRec.LocaleBase.__vfptr.SetLocale(@FComponentBaseRec.LocaleBase, loc);
  finally
  end;
end;

end.
