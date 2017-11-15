unit WrapperObject;

interface

uses
  System.Win.ComServ, System.Win.ComObj, WinApi.ActiveX, System.SysUtils,
  WinApi.Windows, System.Variants, 
  AddInLib, AddInDefBase, MemoryManager, NativeAPILib, ComponentBase, V8Types;

const
  c_AddinName = 'ATOLKKMDriverWrapper';
  CLSID_AddInObject: TGUID = '{90C66C90-7B4F-4C3E-854B-1352117E784F}';

  DriverDLLName = 'fptrwin32_fz54.dll';
  ATOLClassName = 'ATOL_KKT_1C83_V9';

type
  TWrapperObject = class(TComObject, IinitDone, ILanguageExtender)
  private
    F1CIntf: IDispatch;
    //FStatusLine: IStatusLine;
    //FExtWndsSupport: IExtWndsSupport;
    FErrorLog: IErrorLog;
    //FEvent : IAsyncEvent;
    //FProfile : IPropertyProfile;

    FSelfPath: WideString;
    FAddInDefBase: TAddInDefBase;
    FMemoryManager: TMemoryManager;
    FClassObject: TNativeAPIClass;

    FDllHandle: HMODULE;
    FDllInfo: TVSFixedFileInfo;

    function InitATOLObject: Boolean;
    procedure DestroyATOLObject;
    procedure ShowLogString(const LogString: WideString; const MessageType: Integer);
  protected
    { IInitDone implementation }
    function Init(const pConnection: IDispatch): HResult; stdcall;
    function Done: HResult; stdcall;
    function GetInfo(var pInfo: PSafeArray): HResult; stdcall;

    { ILanguageExtender implementation }
    function RegisterExtensionAs(var bstrExtensionName: WideString): HResult; stdcall;
    function GetNProps(var plProps: Integer): HResult; stdcall;
    function FindProp(const bstrPropName: WideString; var plPropNum: Integer): HResult; stdcall;
    function GetPropName(lPropNum: Integer; lPropAlias: Integer; var pbstrPropName: WideString): HResult; stdcall;
    function GetPropVal(lPropNum: Integer; var pvarPropVal: OleVariant): HResult; stdcall;
    function SetPropVal(lPropNum: Integer; var varPropVal: OleVariant): HResult; stdcall;
    function IsPropReadable(lPropNum: Integer; var pboolPropRead: Integer): HResult; stdcall;
    function IsPropWritable(lPropNum: Integer; var pboolPropWrite: Integer): HResult; stdcall;
    function GetNMethods(var plMethods: Integer): HResult; stdcall;
    function FindMethod(const bstrMethodName: WideString; var plMethodNum: Integer): HResult; stdcall;
    function GetMethodName(lMethodNum: Integer; lMethodAlias: Integer; var pbstrMethodName: WideString): HResult; stdcall;
    function GetNParams(lMethodNum: Integer; var plParams: Integer): HResult; stdcall;
    function GetParamDefValue(lMethodNum: Integer; lParamNum: Integer; var pvarParamDefValue: OleVariant): HResult; stdcall;
    function HasRetVal(lMethodNum: Integer; var pboolRetValue: Integer): HResult; stdcall;
    function CallAsProc(lMethodNum: Integer; var paParams: PSafeArray): HResult; stdcall;
    function CallAsFunc(lMethodNum: Integer; var pvarRetValue: OleVariant; var paParams: PSafeArray): HResult; stdcall;
  public
    constructor Create;
    destructor Destroy; override;

  end;

implementation

constructor TWrapperObject.Create;
begin
  inherited;

  FSelfPath := '';
  FDllHandle := 0;
  FAddInDefBase := nil;
  FMemoryManager := nil;
  FClassObject := nil;
end;

destructor TWrapperObject.Destroy;
begin
  inherited Destroy;
end;

function TWrapperObject.InitATOLObject: Boolean;
var
  GetClassObjectFunc: TGetClassObjectFunc;
  DestroyObjectFunc: TDestroyObjectFunc;
  strLen: Integer;
  hHandle: Cardinal;
  pwcInfo: Pointer;
  pffiVersion: PVSFixedFileInfo;
  res: Integer;
begin
  Result := True;
  SetErrorMode(SEM_FAILCRITICALERRORS);

  FSelfPath := ExtractFilePath(GetModuleName(hInstance));

  DestroyObjectFunc := nil;
  FDllInfo.dwFileVersionMS := 0;
  FDllInfo.dwFileVersionLS := 0;

  FClassObject := TNativeAPIClass.Create;
  if (not Assigned(FMemoryManager)) then FMemoryManager := TMemoryManager.Create;
  if (not Assigned(FAddInDefBase)) then FAddInDefBase := TAddInDefBase.Create;

  if not FileExists(FSelfPath + DriverDllName) then
    begin
      ShowLogString('Библиотека ' + DriverDllName + ' не найдена!', ADDIN_E_FAIL);
      Result := False;
    end
  else
    begin
      strLen := GetFileVersionInfoSize(PWideChar(FSelfPath + DriverDLLName), hHandle);
      if strLen > 0 then
        begin
          GetMem(pwcInfo, strLen);
          GetFileVersionInfo(PWideChar(FSelfPath + DriverDLLName), hHandle, strLen, pwcInfo);
          if VerQueryValue(Pointer(pwcInfo), '\', Pointer(pffiVersion), hHandle) then
            FDllInfo := pffiVersion^;
          FreeMem(pwcInfo);
        end;

      if (HiWord(FDllInfo.dwFileVersionMS) < 9) or ((HiWord(FDllInfo.dwFileVersionMS) >= 9) and (LoWord(FDllInfo.dwFileVersionMS) < 11)) then
        begin
          ShowLogString('Версия драйвера ниже 9.11 не поддерживается!', ADDIN_E_FAIL);
          Result := False;
        end
      else
        begin
          FDllHandle := LoadLibrary(PWideChar(FSelfPath + DriverDllName));
          if FDllHandle = 0 then
            begin
              ShowLogString('Ошибка загрузки библиотеки ' + DriverDllName + '!', ADDIN_E_FAIL);
              Result := False;
            end
          else
            begin
              GetClassObjectFunc := GetProcAddress(FDllHandle, 'GetClassObject');
              DestroyObjectFunc := GetProcAddress(FDllHandle, 'DestroyObject');

              if (not Assigned(GetClassObjectFunc)) or (not Assigned(DestroyObjectFunc)) then
                begin
                  ShowLogString('Библиотека ' + DriverDllName + 'не является 1C Native API библиотекой!', ADDIN_E_FAIL);
                  Result := False;
                end
              else
                begin
                  res := 0;
                  try
                    res := GetClassObjectFunc(PWideChar(ATOLClassName), @FClassObject.ComponentBaseRec);
                  except
                  end;
                  
                  if res = 0 then
                    begin
                      ShowLogString('Не удалось создать экземпляр класса ' + ATOLClassName + '!', ADDIN_E_FAIL);
                      Result := False;
                    end
                  else
                    begin
                      try
                        FClassObject.SetMemManager(@FMemoryManager.MemoryManagerRec);
                        if not FClassObject.Init(@FAddInDefBase.AddInDefBaseRec) then
                          begin
                            ShowLogString('Ошибка инициализации класса ' + ATOLClassName + '!', ADDIN_E_FAIL);
                            Result := False;
                          end;
                      except
                        ShowLogString('Ошибка инициализации класса ' + ATOLClassName + '!', ADDIN_E_FAIL);
                        Result := False;
                      end;
                    end;
                end;
            end;
        end;
    end;

  if not Result then
    begin
      if Assigned(FClassObject) then
        begin
          if Assigned(DestroyObjectFunc) then
            try
              DestroyObjectFunc(@FClassObject.ComponentBaseRec);
            except
            end;
          FClassObject.Free;
        end;

      FClassObject := nil;

      if Assigned(FMemoryManager) then FMemoryManager.Free;
      if Assigned(FAddInDefBase) then FAddInDefBase.Free;
      FMemoryManager := nil;
      FAddInDefBase := nil;

      if FDLLHandle <> 0 then FreeLibrary(FDllHandle);
      FDLLHandle := 0;
    end;
end;

procedure TWrapperObject.DestroyATOLObject;
var
  DestroyObjectFunc: TDestroyObjectFunc;
begin
  DestroyObjectFunc := nil;

  if FDLLHandle <> 0 then
    DestroyObjectFunc := GetProcAddress(FDllHandle, 'DestroyObject');

  if Assigned(FClassObject) then
    begin
      try FClassObject.Done() except end;
      if Assigned(DestroyObjectFunc) then
        DestroyObjectFunc(@FClassObject.ComponentBaseRec);
      FClassObject.Free;
    end;

  FClassObject := nil;

  if Assigned(FMemoryManager) then FMemoryManager.Free;
  if Assigned(FAddInDefBase) then FAddInDefBase.Free;
  FMemoryManager := nil;
  FAddInDefBase := nil;

  if FDLLHandle <> 0 then FreeLibrary(FDllHandle);
  FDLLHandle := 0;
end;

procedure TWrapperObject.ShowLogString(const LogString: WideString; const MessageType: Integer);
var
  ErrInfo: EXCEPINFO;
begin
  if not Assigned(FErrorLog) then Exit;
  If Trim(LogString) = '' then Exit;

  ErrInfo.bstrSource := c_AddinName;
  ErrInfo.bstrDescription := LogString;
  ErrInfo.wCode := MessageType;
  ErrInfo.sCode := S_OK;
  try
    FErrorLog.AddError(nil, ErrInfo);
  finally
  end;
end;

function TWrapperObject.Init(const pConnection: IDispatch): HResult; stdcall;
begin
  if not Assigned(pConnection) then
    Result := S_FALSE
  else
    begin
      F1CIntf := pConnection;
      Result := S_OK;

      if Assigned(FErrorLog) then FErrorLog._Release();
      //if Assigned(FEvent) then FEvent._Release();
      //if Assigned(FProfile) then FProfile._Release();
      //if Assigned(FStatusLine) then FStatusLine._Release();
      //if Assigned(FExtWndsSupport) then FExtWndsSupport._Release();

      FErrorLog := nil;
      //FEvent := nil;
      //FProfile := nil;
      //FStatusLine := nil;
      //FExtWndsSupport := nil;

      try
        pConnection.QueryInterface(IID_IErrorLog, FErrorLog);
        //pConnection.QueryInterface(IID_IAsyncEvent, FEvent);
        //pConnection.QueryInterface(IID_IPropertyProfile, FProfile);
        //pConnection.QueryInterface(IID_IStatusLine, FStatusLine);
        //pConnection.QueryInterface(IID_IExtWndsSupport, FExtWndsSupport);
      except
        Result := S_FALSE;
      end;
    end;

  if (Result = S_OK) and (not Assigned(FClassObject)) then
    if not InitATOLObject then Result := S_FALSE;

  if Result = S_FALSE then
    begin
      if Assigned(FErrorLog) then FErrorLog._Release();
      //if Assigned(FEvent) then FEvent._Release();
      //if Assigned(FProfile) then FProfile._Release();
      //if Assigned(FStatusLine) then FStatusLine._Release();
      //if Assigned(FExtWndsSupport) then FExtWndsSupport._Release();
    end;
end;

function TWrapperObject.Done: HResult; stdcall;
begin
  DestroyATOLObject;

  if Assigned(FErrorLog) then FErrorLog._Release();
  //if Assigned(FEvent) then FEvent._Release();
  //if Assigned(FProfile) then FProfile._Release();
  //if Assigned(FStatusLine) then FStatusLine._Release();
  //if Assigned(FExtWndsSupport) then FExtWndsSupport._Release();

  Done := S_OK;
end;

function TWrapperObject.GetInfo(var pInfo: PSafeArray): HResult; stdcall;
var
  varInfo : OleVariant;
  i: Integer;
begin
  varInfo := '2000';
  i := 0;
  SafeArrayPutElement(pInfo, i, varInfo);

  result := S_OK;
end;

function TWrapperObject.RegisterExtensionAs(var bstrExtensionName: WideString): HResult; stdcall;
begin
  bstrExtensionName := c_AddinName;
  result := S_OK;
end;

function TWrapperObject.GetNProps(var plProps: Integer): HResult; stdcall;
begin
  Result := S_OK;
  plProps := 1;
end;

function TWrapperObject.FindProp(const bstrPropName: WideString; var plPropNum: Integer): HResult; stdcall;
begin
  Result := S_FALSE;
  plPropNum := 0;

  if (bstrPropName = 'DriverVersion') or (bstrPropName = 'ВерсияДрайвера') then
    begin
      Result := S_OK;
      plPropNum := 1;
    end;
end;

function TWrapperObject.GetPropName(lPropNum: Integer; lPropAlias: Integer; var pbstrPropName: WideString): HResult; stdcall;
begin
  Result := S_FALSE;
  pbstrPropName := '';

  if lPropNum = 1 then
    begin
      if lPropAlias = 1 then pbstrPropName := 'ВерсияДрайвера' else pbstrPropName := 'DriverVersion';
      Result := S_OK;
    end;
end;

function TWrapperObject.GetPropVal(lPropNum: Integer; var pvarPropVal: OleVariant): HResult; stdcall;
begin
  Result := S_FALSE;
  pvarPropVal := VT_EMPTY;

  if lPropNum = 1 then
    begin
      pvarPropVal := IntToStr(HiWord(FDllInfo.dwFileVersionMS)) + '.' + IntToStr(LoWord(FDllInfo.dwFileVersionMS)) + '.' +
                     IntToStr(HiWord(FDllInfo.dwFileVersionLS)) + '.' + IntToStr(LoWord(FDllInfo.dwFileVersionLS));
      Result := S_OK;
    end;
end;

function TWrapperObject.SetPropVal(lPropNum: Integer; var varPropVal: OleVariant): HResult; stdcall;
begin
  Result := S_FALSE;
end;

function TWrapperObject.IsPropReadable(lPropNum: Integer; var pboolPropRead: Integer): HResult; stdcall;
begin
  Result := S_FALSE;
  pboolPropRead := 0;

  if lPropNum = 1 then
    begin
      pboolPropRead := 1;
      Result := S_OK;
    end;
end;

function TWrapperObject.IsPropWritable(lPropNum: Integer; var pboolPropWrite: Integer): HResult; stdcall;
begin
  Result := S_FALSE;
  pboolPropWrite := 0;
end;

function TWrapperObject.GetNMethods(var plMethods: Integer): HResult; stdcall;
begin
  Result := S_FALSE;
  plMethods := 0;

  if Assigned(FClassObject) then
    try
      plMethods := FClassObject.GetNMethods;
      Result := S_OK;
    except
      ShowLogString('Ошибка при вызове метода GetNMethods!', ADDIN_E_FAIL);
    end
end;

function TWrapperObject.FindMethod(const bstrMethodName: WideString; var plMethodNum: Integer): HResult; stdcall;
begin
  Result := S_FALSE;
  plMethodNum := -1;

  if Assigned(FClassObject) then
    try
      plMethodNum := FClassObject.FindMethod(PWideChar(bstrMethodName));
      Result := S_OK;
    except
      ShowLogString('Ошибка при вызове метода FindMethod!', ADDIN_E_FAIL);
    end;
end;

function TWrapperObject.GetMethodName(lMethodNum: Integer; lMethodAlias: Integer; var pbstrMethodName: WideString): HResult; stdcall;
begin
  Result := S_FALSE;
  pbstrMethodName := '';

  if Assigned(FClassObject) then
    try
      pbstrMethodName := FClassObject.GetMethodName(lMethodNum, lMethodAlias);
      Result := S_OK;
    except
      ShowLogString('Ошибка при вызове метода GetMethodName!', ADDIN_E_FAIL);
    end;
end;

function TWrapperObject.GetNParams(lMethodNum: Integer; var plParams: Integer): HResult; stdcall;
begin
  Result := S_FALSE;
  plParams := 0;

  if Assigned(FClassObject) then
    try
      plParams := FClassObject.GetNParams(lMethodNum);
      Result := S_OK;
    except
      ShowLogString('Ошибка при вызове метода GetNParams!', ADDIN_E_FAIL);
    end;
end;

function TWrapperObject.GetParamDefValue(lMethodNum: Integer; lParamNum: Integer; var pvarParamDefValue: OleVariant): HResult; stdcall;
var
  V8ParamDefValue: T1CVariant;
begin
  Result := S_FALSE;
  pvarParamDefValue := VT_EMPTY;

  if Assigned(FClassObject) then
    try
      if FClassObject.GetParamDefValue(lMethodNum, lParamNum, @V8ParamDefValue) then
        begin
          FromV8Variant(V8ParamDefValue, pvarParamDefValue);
          if V8ParamDefValue.vt = VTYPE_BOOL then
            if V8ParamDefValue.Value.bVal then pvarParamDefValue := 1 else pvarParamDefValue := 0;

          if V8ParamDefValue.vt = VTYPE_PSTR then
            FMemoryManager.FreeMemory(Pointer(V8ParamDefValue.Value.pstrVal));
          if V8ParamDefValue.vt = VTYPE_PWSTR then
            FMemoryManager.FreeMemory(Pointer(V8ParamDefValue.Value.pwstrVal));
        end;
    except
      ShowLogString('Ошибка при вызове метода GetParamDefValue!', ADDIN_E_FAIL);
    end;
end;

function TWrapperObject.HasRetVal(lMethodNum: Integer; var pboolRetValue: Integer): HResult; stdcall;
begin
  Result := S_FALSE;
  pboolRetValue := 0;

  if Assigned(FClassObject) then
    try
      if FClassObject.HasRetVal(lMethodNum) then
        pboolRetValue := 1;
      Result := S_OK;
    except
      ShowLogString('Ошибка при вызове метода HasRetVal!', ADDIN_E_FAIL);
    end;
end;

function TWrapperObject.CallAsProc(lMethodNum: Integer; var paParams: PSafeArray): HResult; stdcall;
var
  V8Params: array of T1CVariant;
  ParamsCount: Integer;
  i: Integer;
  tParam: OleVariant;
  MethodName: WideString;
begin
  Result := S_FALSE;

  ParamsCount := 0;
  GetNParams(lMethodNum, ParamsCount);

  if Assigned(FClassObject) then
    try
      MethodName := FClassObject.GetMethodName(lMethodNum, 0);

      SetLength(V8Params, ParamsCount);
      for i := 0 to ParamsCount - 1 do
        begin
          try
            SafeArrayGetElement(paParams, i, tParam);
            ToV8Variant(tParam, V8Params[i], FMemoryManager);

            if (MethodName = 'ProcessCheck') and (i = 1) then
              begin
                V8Params[i].vt := VTYPE_BOOL;
                V8Params[i].Value.bVal := tParam <> 0;
              end;
          except
            V8Params[i].vt := VTYPE_EMPTY;
          end;
        end;

      if FClassObject.CallAsProc(lMethodNum, Pointer(V8Params), ParamsCount) then
        begin
          for i := 0 to ParamsCount - 1 do
            begin
              FromV8Variant(V8Params[i], tParam);
              if V8Params[i].vt = VTYPE_BOOL then
                if V8Params[i].Value.bVal then tParam := 1 else tParam := 0;
              try SafeArrayPutElement(paParams, i, tParam) except end;
            end;
          Result := S_OK;
        end;

      for i := 0 to ParamsCount - 1 do
        begin
          if V8Params[i].vt = VTYPE_PSTR then
            FMemoryManager.FreeMemory(Pointer(V8Params[i].Value.pstrVal));
          if V8Params[i].vt = VTYPE_PWSTR then
            FMemoryManager.FreeMemory(Pointer(V8Params[i].Value.pwstrVal));
        end;
    except
      ShowLogString('Ошибка при вызове метода CallAsProc!', ADDIN_E_FAIL);
    end;

  SetLength(V8Params, 0);
end;

function TWrapperObject.CallAsFunc(lMethodNum: Integer; var pvarRetValue: OleVariant; var paParams: PSafeArray): HResult; stdcall;
var
  tParam: OleVariant;
  V8Params: array of T1CVariant;
  V8RetValue: T1CVariant;
  ParamsCount: Integer;
  i: Integer;
  MethodName: WideString;
begin
  Result := S_FALSE;

  ParamsCount := 0;
  GetNParams(lMethodNum, ParamsCount);

  if Assigned(FClassObject) then
    try
      MethodName := FClassObject.GetMethodName(lMethodNum, 0);

      SetLength(V8Params, ParamsCount);
      for i := 0 to ParamsCount - 1 do
        begin
          try
            SafeArrayGetElement(paParams, i, tParam);
            ToV8Variant(tParam, V8Params[i], FMemoryManager);

            if (MethodName = 'ProcessCheck') and (i = 1) then
              begin
                V8Params[i].vt := VTYPE_BOOL;
                V8Params[i].Value.bVal := tParam <> 0;
              end;
          except
            V8Params[i].vt := VTYPE_EMPTY;
          end;
        end;
      V8RetValue.vt := VTYPE_EMPTY;

      if FClassObject.CallAsFunc(lMethodNum, @V8RetValue, Pointer(V8Params), ParamsCount) then
        begin
          for i := 0 to ParamsCount - 1 do
            begin
              FromV8Variant(V8Params[i], tParam);
              if V8Params[i].vt = VTYPE_BOOL then
                if V8Params[i].Value.bVal then tParam := 1 else tParam := 0;
              try SafeArrayPutElement(paParams, i, tParam) except end;
            end;
          FromV8Variant(V8RetValue, pvarRetValue);
          if V8RetValue.vt = VTYPE_BOOL then
            if V8RetValue.Value.bVal then pvarRetValue := 1 else pvarRetValue := 0;
          Result := S_OK;
        end;

      for i := 0 to ParamsCount - 1 do
        begin
          if V8Params[i].vt = VTYPE_PSTR then
            FMemoryManager.FreeMemory(Pointer(V8Params[i].Value.pstrVal));
          if V8Params[i].vt = VTYPE_PWSTR then
            FMemoryManager.FreeMemory(Pointer(V8Params[i].Value.pwstrVal));
        end;
    except
      ShowLogString('Ошибка при вызове метода CallAsProc!', ADDIN_E_FAIL);
    end;

  SetLength(V8Params, 0);
end;

initialization
  ComServer.SetServerName('AddIn');
  TComObjectFactory.Create(ComServer, TWrapperObject, CLSID_AddInObject, c_AddinName, 'V7 AddIn 2.0', ciMultiInstance);

end.
