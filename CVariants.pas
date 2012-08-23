unit CVariants;

interface

uses
  Collections;

{$INCLUDE 'CVariantDelphiFeatures.inc'}

{$WARN UNSAFE_TYPE OFF}
{$WARN UNSAFE_CODE OFF}
{$WARN UNSAFE_CAST OFF}

const
  vtList  = Collections.vtList;
  vtMap   = Collections.vtMap;
  vtEmpty = Collections.vtEmpty;
  vtNull  = Collections.vtNull;

type
  {$IFNDEF DELPHI_IS_UNICODE}
  UnicodeString = type WideString;
  {$ENDIF}

  // CVariant must have the same size as Variant
  // There must be Variant inside and nothing else
  // CVariant stands for collection-variant
  CVariant = packed {$IFNDEF DELPHI_HAS_RECORDS} object {$ELSE} record {$ENDIF}
  private
    FObj: Variant;

    constructor CreateI(const Int: IUnknown; NilToNull: Boolean = True);
    class function VariantToRef(const Obj: Variant): IUnknown;
    class function ConstToRef(const Obj: TVarRec): IUnknown;
    class function GetTVarRecType(const Obj: Variant): SmallInt;
    // class function MakeI(const Int: IUnknown; NilToNull: Boolean = True): CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    function GetAsPVariant: PVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    function GetHash: Integer;

    function GetVType: SmallInt;

    // maps and lists
    function GetItems(const Ind: Variant): CVariant;
    procedure SetItems(const Ind: Variant; const NewObj: CVariant);

    procedure RaiseNotAnArray;
    function GetCollection: IUnknown;
    function GetDeepItem(const Indices: array of const): IUnknown;
    function GetDeepParent(const Indices: array of const): IUnknown;
    function GetDeepParent2(const IndicesAndObj: array of const): IUnknown;

    function GetIsEmpty: Boolean;
    function GetIsFull: Boolean;
    function GetSize: Integer;
  public
    destructor Destroy; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor CreateNull; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor CreateOwned(Obj: TObject); {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor Create(Obj: TObject); overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor Create(const Str: string);  overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor Create(Int: Integer); overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor Create(Dbl: Double);  overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor Create(Bol: Boolean); overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor CreateV(const Vrn: Variant); {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor CreateL; overload;
    constructor CreateL(const AItems: array of const); overload;
    constructor CreateM; overload;
    constructor CreateM(const AKeyValues: array of const); overload;
    constructor CreateM(const AKeys, AValues: array of const); overload;
    function ToString: string;
    function ToInt: Integer;
    function ToBool: Boolean;

    // maps and lists

    function Get(const Indices: array of const): CVariant; // I failed to make a property out of them
    function Has(const Indices: array of const): Boolean;
    procedure Put(const Indices: array of const; const NewObj: CVariant); overload;
    procedure Put(const IndicesAndObj: array of const); overload;
    procedure Remove(const Indices: array of const);
    procedure Insert(const Indices: array of const; const NewObj: CVariant); overload;
    procedure Insert(const IndicesAndObj: array of const); overload;
    procedure Append(const Indices: array of const; const NewObj: CVariant); overload;
    procedure Append(const IndicesAndObj: array of const); overload;
    procedure Append(const NewObj: CVariant); overload;

    procedure Clear(const Indices: array of const); overload; // makes collection empty without destroying
    procedure Clear; overload; // makes collection empty without destroying
    function IsEmptyDeep(const Indices: array of const): Boolean;
    function IsFullDeep(const Indices: array of const): Boolean;
    function SizeDeep(const Indices: array of const): Integer;
    function VTypeDeep(const Indices: array of const): SmallInt;

    function HasKey(const Key: string): Boolean;

    property AsVariant: Variant read FObj write FObj;
    property Items[const Ind: Variant]: CVariant read GetItems write SetItems; // TODO: Delphi 2006 - default?
    property AsPVariant: PVariant read GetAsPVariant;
    property Hash: Integer read GetHash;
    property IsEmpty: Boolean read GetIsEmpty;
    property IsFull: Boolean read GetIsFull;
    property Size: Integer read GetSize;
    property VType: SmallInt read GetVType;
  end;

  CListIterator = {$IFNDEF DELPHI_HAS_RECORDS} object {$ELSE} record {$ENDIF}
  private
    FList, FIterator: IUnknown;
    FPosition: Integer;
    function GetList: CVariant;
  public
    Key: Integer;
    Value: CVariant;
    constructor Create(const AList: CVariant);
    destructor Destroy;
    function Next: Boolean;
    property List: CVariant read GetList;
  end;

  CMapIterator = {$IFNDEF DELPHI_HAS_RECORDS} object {$ELSE} record {$ENDIF}
  private
    FMap, FIterator: IUnknown;
    function GetMap: CVariant;
  public
    Key: string;
    Value: CVariant;
    constructor Create(const AMap: CVariant);
    destructor Destroy;
    function Next: Boolean;
    property Map: CVariant read GetMap;
  end;

function CVarOwned(Obj: TObject): CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVarEmpty: CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVarNull: CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVar(Obj: TObject): CVariant; overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVar(const Str: string): CVariant;  overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVar(Int: Integer): CVariant; overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVar(Dbl: Double): CVariant;  overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVar(Bol: Boolean): CVariant; overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVarV(const Vrn: Variant): CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CList: CVariant; overload;
function CList(const AItems: array of const): CVariant; overload;
function CMap: CVariant; overload;
function CMap(const AKeyValues: array of const): CVariant; overload;
function CMap(const AKeys, AValues: array of const): CVariant; overload;

function VList: Variant; overload;
function VList(const AItems: array of const): Variant; overload;
function VMap: Variant; overload;
function VMap(const AKeyValues: array of const): Variant; overload;
function VMap(const AKeys, AValues: array of const): Variant; overload;

implementation

uses
  SysUtils, Variants, Math;

class function CVariant.VariantToRef(const Obj: Variant): IUnknown;
begin
  Result := nil;
  case TVarData(Obj).VType of
    varEmpty: Result := iempty;
    varNull: Exit;
    varSmallint: Result := iref(TVarData(Obj).VSmallInt);
    varInteger: Result := iref(TVarData(Obj).VInteger);
    varSingle: Result := iref(TVarData(Obj).VSingle);
    varDouble: Result := iref(TVarData(Obj).VDouble);
    varCurrency: Result := iref(TVarData(Obj).VCurrency);
    varDate: Result := iref(TVarData(Obj).VDate);
    varOleStr: Result := iref(WideString(Pointer(TVarData(Obj).VOleStr)));
    varDispatch: Result := IUnknown(TVarData(Obj).VDispatch);
    varError: Result := iref(TVarData(Obj).VError);
    varBoolean: Result := iref(TVarData(Obj).VBoolean);
    // varVariant (only references are possible)
    varUnknown: Result := IUnknown(TVarData(Obj).VUnknown);
    varShortInt: Result := iref(TVarData(Obj).VShortInt);
    varByte: Result := iref(TVarData(Obj).VByte);
    varWord: Result := iref(TVarData(Obj).VWord);
    varLongWord: Result := iref(TVarData(Obj).VLongWord);
    varInt64: Result := iref(TVarData(Obj).VInt64);
    varString: Result := iref(AnsiString(Pointer(TVarData(Obj).VString)));
    varArray: // TODO: COM array to Collection sequence
  else
    // TODO: custom Variant
    // TODO: Delphi XE2 types
    // TODO: pointer to something
  end;
end;

class function CVariant.GetTVarRecType(const Obj: Variant): SmallInt;
var
  LIU: IUnknown;
begin
  Result := vtVariant;
  case TVarData(Obj).VType of
    varEmpty: Result := vtEmpty;
    varNull: Result := vtNull;
    varSmallint: Result := vtInteger;
    varInteger: Result := vtInteger;
    varSingle: Result := vtExtended;
    varDouble: Result := vtExtended;
    varCurrency: Result := vtExtended;
    varDate: Result := vtExtended;
    varOleStr: Result := vtString;
    varDispatch: begin
      Result := vtInterface;
      LIU := IDispatch(TVarData(Obj).VDispatch);
    end;
    varError: Result := vtInteger;
    varBoolean: Result := vtBoolean;
    // varVariant (only references are possible)
    varUnknown: begin
      Result := vtInterface;
      LIU := IUnknown(TVarData(Obj).VUnknown);
    end;
    varShortInt: Result := vtInteger;
    varByte: Result := vtInteger;
    varWord: Result := vtInteger;
    varLongWord: Result := vtInteger;
    varInt64: Result := vtInt64;
    varString: Result := vtString;
    varArray: // TODO: COM array to Collection sequence
  else
    // TODO: custom Variant
    // TODO: Delphi XE2 types
    // TODO: pointer to something
  end;
  if Result = vtInterface then
  begin
    if not Assigned(LIU) then
      Result := vtEmpty
    else if Supports(LIU, IList) then
      Result := vtList
    else if Supports(LIU, IMap) then
      Result := vtMap;
  end;
end;

function CVariant.GetVType: SmallInt;
begin
  Result := GetTVarRecType(FObj);
end;

function CVariant.VTypeDeep(const Indices: array of const): SmallInt;
var
  LIU: IUnknown;
  LIO: IObject;
begin
  if Length(Indices) <= 0 then
    Result := GetVType
  else
  try
    LIU := GetDeepItem(Indices);
    Result := vtInterface;

    if not Assigned(LIU) then
      Result := vtEmpty
    else if Supports(LIU, IObject, LIO) then
      Result := LIO.VType;
  except
    on EVariantNotAnArrayError do
      Result := vtNull;
    on EVariantInvalidArgError do
      Result := vtNull;
    on EVariantBadIndexError do
      Result := vtNull;
  end;
end;

class function CVariant.ConstToRef(const Obj: TVarRec): IUnknown;
begin
  Result := nil;
  case Obj.VType of
    vtInteger: Result := iref(Obj.VInteger);
    vtBoolean: Result := iref(Obj.VBoolean);
    vtChar: Result := iref(Obj.VChar);
    vtExtended: Result := iref(Obj.VExtended^);
    vtString: Result := iref(Obj.VString^);
    vtPChar: Result := iref(AnsiString(Obj.VPChar));
    vtObject: Result := iref(Obj.VObject);
    vtWideChar: Result := iref(Obj.VWideChar);
    vtPWideChar: Result := iref(Obj.VPWideChar);
    vtAnsiString: Result := iref(AnsiString(Obj.VAnsiString));
    vtCurrency: Result := iref(Obj.VCurrency^);
    vtVariant: Result := VariantToRef(Obj.VVariant^);
    vtInterface:
      if Assigned(Obj.VInterface) then
        Result := IUnknown(Obj.VInterface)
      else
        Result := iempty;
    vtWideString: Result := iref(WideString(Obj.VWideString));
    vtInt64: Result := iref(Obj.VInt64^);
    vtPointer: if Obj.VPointer = nil then Result := iempty;
  else
    // TODO: Delphi XE2 types
  end;
end;

constructor CVariant.CreateI(const Int: IUnknown; NilToNull: Boolean = True);
var
  IntObj: IObject;
begin
  if not Assigned(Int) then
  begin
    if NilToNull then
      CreateNull
    else
      Destroy;
  end else
  if Supports(Int, IObject, IntObj) then
  begin
    case IntObj.VType of
    vtEmpty: Destroy;
    vtInteger: FObj := (IntObj as IInteger).intValue;
    vtExtended: FObj := (IntObj as IDouble).doubleValue;
    vtBoolean: FObj := (IntObj as IBoolean).boolValue;
    vtString: FObj := (IntObj as IString).toString;
    else
      FObj := Int;
    end;
  end else
  begin
    FObj := Int;
  end;
end;

// class function CVariant.MakeI(const Int: IUnknown; NilToNull: Boolean = True): CVariant;
// begin
//   Result.CreateI(Int, NilToNull);
// end;

destructor CVariant.Destroy;
begin
  FObj := Unassigned;
end;

constructor CVariant.CreateNull;
begin
  FObj := Null;
end;

constructor CVariant.CreateOwned(Obj: TObject);
begin
  FObj := iown(Obj);
end;

constructor CVariant.Create(Obj: TObject);
begin
  FObj := iref(Obj);
end;

constructor CVariant.Create(const Str: string);
begin
  FObj := Str;
end;

constructor CVariant.Create(Int: Integer);
begin
  FObj := Int;
end;

constructor CVariant.Create(Dbl: Double);
begin
  FObj := Dbl;
end;

constructor CVariant.Create(Bol: Boolean);
begin
  FObj := Bol;
end;

constructor CVariant.CreateV(const Vrn: Variant);
begin
  FObj := Vrn;
end;

constructor CVariant.CreateL;
begin
  CreateI(TArrayList.create);
end;

constructor CVariant.CreateL(const AItems: array of const);
var
  LIL: IList;
  i: Integer;
begin
  LIL := TArrayList.create(Length(AItems));
  CreateI(LIL);
  for i := Low(AItems) to High(AItems) do
  begin
    LIL.add(ConstToRef(AItems[i]));
  end;
end;

constructor CVariant.CreateM;
begin
  CreateI(THashMap.create);
end;

constructor CVariant.CreateM(const AKeyValues: array of const);
var
  LIM: IMap;
  i: Integer;
begin
  LIM := THashMap.create;
  CreateI(LIM);
  for i := 0 to Length(AKeyValues) div 2 - 1 do
  begin
    LIM.put(ConstToRef(AKeyValues[Low(AKeyValues) + i * 2]),
            ConstToRef(AKeyValues[Low(AKeyValues) + i * 2 + 1]));
  end;
end;

constructor CVariant.CreateM(const AKeys, AValues: array of const);
var
  LIM: IMap;
  i: Integer;
begin
  LIM := THashMap.create;
  CreateI(LIM);
  for i := 0 to Min(Length(AKeys), Length(AValues)) - 1 do
  begin
    LIM.put(ConstToRef(AKeys[Low(AKeys) + i]),
            ConstToRef(AValues[Low(AValues) + i]));
  end;
end;

function CVarOwned(Obj: TObject): CVariant;
begin
  Result.FObj := iown(Obj);
end;

function CVarEmpty: CVariant;
begin
  Result.FObj := Unassigned;
end;

function CVarNull: CVariant;
begin
  Result.FObj := Null;
end;

function CVar(Obj: TObject): CVariant;
begin
  Result.FObj := iref(Obj);
end;

function CVar(const Str: string): CVariant;
begin
  Result.FObj := Str;
end;

function CVar(Int: Integer): CVariant;
begin
  Result.FObj := Int;
end;

function CVar(Dbl: Double): CVariant;
begin
  Result.FObj := Dbl;
end;

function CVar(Bol: Boolean): CVariant;
begin
  Result.FObj := Bol;
end;

function CVarV(const Vrn: Variant): CVariant;
begin
  Result.FObj := Vrn;
end;

function CList: CVariant;
begin
  Result.CreateL;
end;

function CList(const AItems: array of const): CVariant;
begin
  Result.CreateL(AItems);
end;

function CMap: CVariant;
begin
  Result.CreateM;
end;

function CMap(const AKeyValues: array of const): CVariant;
begin
  Result.CreateM(AKeyValues);
end;

function CMap(const AKeys, AValues: array of const): CVariant;
begin
  Result.CreateM(AKeys, AValues);
end;

function VList: Variant;
begin
  CVariant(Result).CreateL;
end;

function VList(const AItems: array of const): Variant;
begin
  CVariant(Result).CreateL(AItems);
end;

function VMap: Variant;
begin
  CVariant(Result).CreateM;
end;

function VMap(const AKeyValues: array of const): Variant;
begin
  CVariant(Result).CreateM(AKeyValues);
end;

function VMap(const AKeys, AValues: array of const): Variant;
begin
  CVariant(Result).CreateM(AKeys, AValues);
end;

function CVariant.GetAsPVariant: PVariant;
begin
  Result := @FObj;
end;

function CVariant.GetHash: Integer;
begin
  Result := hashOf(VariantToRef(FObj));
end;

function CVariant.ToString: string;
begin
  Result := stringOf(VariantToRef(FObj));
end;

function CVariant.ToInt: Integer;
begin
  Result := intOf(VariantToRef(FObj));
end;

function CVariant.ToBool: Boolean;
begin
  Result := boolOf(VariantToRef(FObj));
end;

procedure CVariant.RaiseNotAnArray;
begin
  raise EVariantNotAnArrayError.Create('Not an array being indexed');
end;

function CVariant.GetCollection: IUnknown;
begin
  Result := nil;
  case TVarData(FObj).VType of
  varUnknown: Result := IUnknown(TVarData(FObj).VUnknown);
  varDispatch: Result := IDispatch(TVarData(FObj).VDispatch);
  else
    RaiseNotAnArray;
  end;
end;

function CVariant.GetDeepItem(const Indices: array of const): IUnknown;
var
  i: Integer;
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  Result := nil;
  LIU := GetCollection;

  for i := Low(Indices) to High(Indices) do
  begin
    if Supports(LIU, IList, LIL) then
    begin
      if Indices[i].VType <> vtInteger then
        raise EVariantInvalidArgError.Create(stringOf(LIL));
      LIU := LIL.item[Indices[i].VInteger];
    end else if Supports(LIU, IMap, LIM) then
      LIU := LIM.get(ConstToRef(Indices[i]))
    else
      RaiseNotAnArray;
    LIL := nil; LIM := nil;
  end;

  Result := LIU;
end;

function CVariant.GetDeepParent(const Indices: array of const): IUnknown;
var
  i: Integer;
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  Result := nil;
  LIU := GetCollection;

  if Length(Indices) < 1 then
    raise EVariantInvalidArgError.Create('At least one index is expected');

  for i := Low(Indices) to High(Indices) - 1 do
  begin
    if Supports(LIU, IList, LIL) then
    begin
      if Indices[i].VType <> vtInteger then
        raise EVariantInvalidArgError.Create(stringOf(LIL));
      LIU := LIL.item[Indices[i].VInteger];
    end else if Supports(LIU, IMap, LIM) then
      LIU := LIM.get(ConstToRef(Indices[i]))
    else
      RaiseNotAnArray;
    LIL := nil; LIM := nil;
  end;

  Result := LIU;
end;

function CVariant.GetDeepParent2(const IndicesAndObj: array of const): IUnknown;
var
  i: Integer;
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  Result := nil;
  LIU := GetCollection;

  if Length(IndicesAndObj) < 2 then
    raise EVariantInvalidArgError.Create('At least one index and one object is expected');

  for i := Low(IndicesAndObj) to High(IndicesAndObj) - 2 do
  begin
    if Supports(LIU, IList, LIL) then
    begin
      if IndicesAndObj[i].VType <> vtInteger then
        raise EVariantInvalidArgError.Create(stringOf(LIL));
      LIU := LIL.item[IndicesAndObj[i].VInteger];
    end else if Supports(LIU, IMap, LIM) then
      LIU := LIM.get(ConstToRef(IndicesAndObj[i]))
    else
      RaiseNotAnArray;
    LIL := nil; LIM := nil;
  end;

  Result := LIU;
end;

function CVariant.Get(const Indices: array of const): CVariant;
begin
  try
    Result.CreateI(GetDeepItem(Indices), False);
  except
    on EVariantNotAnArrayError do
      Result.CreateNull;
    on EVariantInvalidArgError do
      Result.CreateNull;
    on EVariantBadIndexError do
      Result.CreateNull;
  end;
end;

function CVariant.Has(const Indices: array of const): Boolean;
var
  Dummy: IUnknown;
begin
  try
    Dummy := GetDeepItem(Indices);
    Result := Dummy <> nil;
  except
    on EVariantNotAnArrayError do
      Result := False;
    on EVariantInvalidArgError do
      Result := False;
    on EVariantBadIndexError do
      Result := False;
  end;
end;

procedure CVariant.Put(const Indices: array of const; const NewObj: CVariant);
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  LIU := GetDeepParent(Indices);

  if Supports(LIU, IList, LIL) then
  begin
    if Indices[High(Indices)].VType <> vtInteger then
      raise EVariantInvalidArgError.Create(stringOf(LIL));
    LIL.item[Indices[High(Indices)].VInteger] := VariantToRef(NewObj.FObj);
  end else if Supports(LIU, IMap, LIM) then
    LIM.put(ConstToRef(Indices[High(Indices)]), VariantToRef(NewObj.FObj))
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Put(const IndicesAndObj: array of const);
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  LIU := GetDeepParent2(IndicesAndObj);

  if Supports(LIU, IList, LIL) then
  begin
    if IndicesAndObj[High(IndicesAndObj) - 1].VType <> vtInteger then
      raise EVariantInvalidArgError.Create(stringOf(LIL));
    LIL.item[IndicesAndObj[High(IndicesAndObj) - 1].VInteger] :=
      ConstToRef(IndicesAndObj[High(IndicesAndObj)]);
  end else if Supports(LIU, IMap, LIM) then
    LIM.put(ConstToRef(IndicesAndObj[High(IndicesAndObj) - 1]),
      ConstToRef(IndicesAndObj[High(IndicesAndObj)]))
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Remove(const Indices: array of const);
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  LIU := GetDeepParent(Indices);

  if Supports(LIU, IList, LIL) then
  begin
    if Indices[High(Indices)].VType <> vtInteger then
      raise EVariantInvalidArgError.Create(stringOf(LIL));
    LIL.remove(Indices[High(Indices)].VInteger);
  end else if Supports(LIU, IMap, LIM) then
    LIM.remove(ConstToRef(Indices[High(Indices)]))
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Insert(const Indices: array of const; const NewObj: CVariant);
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  LIU := GetDeepParent(Indices);

  if Supports(LIU, IList, LIL) then
  begin
    if Indices[High(Indices)].VType <> vtInteger then
      raise EVariantInvalidArgError.Create(stringOf(LIL));
    LIL.insert(Indices[High(Indices)].VInteger, VariantToRef(NewObj.FObj));
  end else if Supports(LIU, IMap, LIM) then
    LIM.put(ConstToRef(Indices[High(Indices)]), VariantToRef(NewObj.FObj))
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Insert(const IndicesAndObj: array of const);
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  LIU := GetDeepParent2(IndicesAndObj);

  if Supports(LIU, IList, LIL) then
  begin
    if IndicesAndObj[High(IndicesAndObj) - 1].VType <> vtInteger then
      raise EVariantInvalidArgError.Create(stringOf(LIL));
    LIL.insert(IndicesAndObj[High(IndicesAndObj) - 1].VInteger,
      ConstToRef(IndicesAndObj[High(IndicesAndObj)]));
  end else if Supports(LIU, IMap, LIM) then
    LIM.put(ConstToRef(IndicesAndObj[High(IndicesAndObj) - 1]),
      ConstToRef(IndicesAndObj[High(IndicesAndObj)]))
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Append(const Indices: array of const; const NewObj: CVariant);
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  LIU := GetDeepItem(Indices);

  if Supports(LIU, IList, LIL) then
    LIL.add(VariantToRef(NewObj.FObj))
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Append(const IndicesAndObj: array of const);
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  LIU := GetDeepParent(IndicesAndObj);

  if Supports(LIU, IList, LIL) then
    LIL.add(ConstToRef(IndicesAndObj[High(IndicesAndObj)]))
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Append(const NewObj: CVariant);
var
  LIU: IUnknown;
  LIL: IList;
begin
  LIU := GetCollection;

  if Supports(LIU, IList, LIL) then
    LIL.add(VariantToRef(NewObj.FObj))
  else
    RaiseNotAnArray;
  LIL := nil;
end;

function CVariant.GetItems(const Ind: Variant): CVariant;
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  try
    LIU := GetCollection;

    if Supports(LIU, IList, LIL) then
      LIU := LIL.item[Ind]
    else if Supports(LIU, IMap, LIM) then
      LIU := LIM.get(VariantToRef(Ind))
    else
      RaiseNotAnArray;
    LIL := nil; LIM := nil;

    Result.CreateI(LIU, False);
  except
    on EVariantNotAnArrayError do
      Result.CreateNull;
    on EVariantInvalidArgError do
      Result.CreateNull;
    on EVariantBadIndexError do
      Result.CreateNull;
  end;
end;

procedure CVariant.SetItems(const Ind: Variant; const NewObj: CVariant);
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  LIU := GetCollection;

  if Supports(LIU, IList, LIL) then
    LIL.item[Ind] := VariantToRef(NewObj.FObj)
  else if Supports(LIU, IMap, LIM) then
    LIM.put(VariantToRef(Ind), VariantToRef(NewObj.FObj))
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

function CVariant.GetIsEmpty: Boolean;
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  Result := False;
  LIU := GetCollection;

  if Supports(LIU, IList, LIL) then
    Result := LIL.isEmpty
  else if Supports(LIU, IMap, LIM) then
    Result := LIM.isEmpty
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

function CVariant.GetIsFull: Boolean;
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  Result := False;
  LIU := GetCollection;

  if Supports(LIU, IList, LIL) then
    Result := LIL.isFull
  else if Supports(LIU, IMap, LIM) then
    Result := LIM.isFull
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

function CVariant.GetSize: Integer;
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  Result := 0;
  LIU := GetCollection;

  if Supports(LIU, IList, LIL) then
    Result := LIL.size
  else if Supports(LIU, IMap, LIM) then
    Result := LIM.size
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

function CVariant.IsEmptyDeep(const Indices: array of const): Boolean;
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  Result := False;
  LIU := GetDeepItem(Indices);

  if Supports(LIU, IList, LIL) then
    Result := LIL.isEmpty
  else if Supports(LIU, IMap, LIM) then
    Result := LIM.isEmpty
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

function CVariant.IsFullDeep(const Indices: array of const): Boolean;
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  Result := False;
  LIU := GetDeepItem(Indices);

  if Supports(LIU, IList, LIL) then
    Result := LIL.isFull
  else if Supports(LIU, IMap, LIM) then
    Result := LIM.isFull
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

function CVariant.SizeDeep(const Indices: array of const): Integer;
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  Result := 0;
  LIU := GetDeepItem(Indices);

  if Supports(LIU, IList, LIL) then
    Result := LIL.size
  else if Supports(LIU, IMap, LIM) then
    Result := LIM.size
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Clear(const Indices: array of const);
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  LIU := GetDeepItem(Indices);

  if Supports(LIU, IList, LIL) then
    LIL.clear
  else if Supports(LIU, IMap, LIM) then
    LIM.clear
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Clear;
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  LIU := GetCollection;

  if Supports(LIU, IList, LIL) then
    LIL.clear
  else if Supports(LIU, IMap, LIM) then
    LIM.clear
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;
end;

function CVariant.HasKey(const Key: string): Boolean;
var
  LIU: IUnknown;
  LIM: IMap;
begin
  Result := False;
  try
    LIU := GetCollection;

    if Supports(LIU, IMap, LIM) then
      Result := LIM.has(Key)
    else
      RaiseNotAnArray;
    LIM := nil;
  except
    on EVariantNotAnArrayError do
      Result := False;
    on EVariantInvalidArgError do
      Result := False;
    on EVariantBadIndexError do
      Result := False;
  end;
end;

constructor CMapIterator.Create(const AMap: CVariant);
var
  LIU: IUnknown;
  LIM: IMap;
begin
  LIU := AMap.GetCollection;

  if Supports(LIU, IMap, LIM) then
  begin
    FIterator := LIM.keys;
    FMap := LIM;
    Key := '';
    Value.Destroy;
  end else
    AMap.RaiseNotAnArray;
  LIM := nil;
end;

destructor CMapIterator.Destroy;
begin
  FIterator := nil;
  FMap := nil;
  Key := '';
  Value.Destroy;
end;

function CMapIterator.Next: Boolean;
var
  Next_Key: string;
begin
  Result := False;
  if not Assigned(FIterator) then Exit;
  if IIterator(FIterator).hasNext then
  begin
    Next_Key := IIterator(FIterator).nextStr;
    Key := Next_Key;
    Value.CreateI(IMap(FMap).Get(Next_Key), False);
    Result := True;
  end else
    FIterator := nil;
end;

function CMapIterator.GetMap: CVariant;
begin
  Result.CreateI(FMap);
end;

constructor CListIterator.Create(const AList: CVariant);
var
  LIU: IUnknown;
  LIL: IList;
begin
  LIU := AList.GetCollection;

  if Supports(LIU, IList, LIL) then
  begin
    FIterator := LIL.iterator;
    FList := LIL;
    FPosition := 0;
    Key := 0; Value.Destroy;
  end else
    AList.RaiseNotAnArray;
  LIL := nil;
end;

destructor CListIterator.Destroy;
begin
  FList := nil; FIterator := nil; FPosition := 0;
  Key := 0; Value.Destroy;
end;

function CListIterator.Next: Boolean;
begin
  Result := False;
  if not Assigned(FIterator) then Exit;
  if IIterator(FIterator).hasNext then
  begin
    Key := FPosition;
    Inc(FPosition);
    Value.CreateI(IIterator(FIterator).next, False);
    Result := True;
  end else
    Destroy;
end;

function CListIterator.GetList: CVariant;
begin
  Result.CreateI(FList);
end;

end.
