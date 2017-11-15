unit MemoryManager;

interface

type
  { IMemoryManager }
  TAllocMemoryFunc = function(var pMemory: Pointer; ulCountByte: Cardinal): Boolean of object; stdcall;
  TFreeMemoryFunc = procedure(var pMemory: Pointer) of object; stdcall;

  P1CMemoryManager = ^T1CMemoryManager;
  P1CMemoryManagerVTable = ^T1CMemoryManagerVTable;
  T1CMemoryManager = packed record
    __vfptr: P1CMemoryManagerVTable;
  end;

  T1CMemoryManagerVTable = packed record
    _Destructor: procedure(This: P1CMemoryManager); cdecl;
    AllocMemory: TAllocMemoryFunc;
    FreeMemory: TFreeMemoryFunc;
  end;

  TMemoryManager = class
  private
    FMemoryManagerRec: T1CMemoryManager;
  public
    constructor Create();
    destructor Destroy(); override;
    function AllocMemory(var pMemory: Pointer; ulCountByte: Cardinal): Boolean; stdcall;
    procedure FreeMemory(var pMemory: Pointer); stdcall;

    property MemoryManagerRec: T1CMemoryManager read FMemoryManagerRec;
  end;

implementation

constructor TMemoryManager.Create;
begin
  inherited;

  GetMem(FMemoryManagerRec.__vfptr, SizeOf(T1CMemoryManagerVTable));
  FMemoryManagerRec.__vfptr._Destructor := nil;
  FMemoryManagerRec.__vfptr.AllocMemory := AllocMemory;
  FMemoryManagerRec.__vfptr.FreeMemory := FreeMemory;
end;

destructor TMemoryManager.Destroy;
begin
  FreeMem(FMemoryManagerRec.__vfptr);

  inherited Destroy;
end;

function TMemoryManager.AllocMemory(var pMemory: Pointer; ulCountByte: Cardinal): Boolean; stdcall;
begin
  try
    GetMem(pMemory, ulCountByte);
    FillChar(pMemory^, ulCountByte, 0);
    Result := true;

  except
    Result := false;
  end;
end;

procedure TMemoryManager.FreeMemory(var pMemory: Pointer); stdcall;
begin
  try
    FreeMem(pMemory);
  finally
  end;
end;

end.
