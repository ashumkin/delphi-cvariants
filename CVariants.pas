unit CVariants;

interface

{$INCLUDE 'CVariantDelphiFeatures.inc'}

{$WARN UNSAFE_TYPE OFF}
{$WARN UNSAFE_CODE OFF}
{$WARN UNSAFE_CAST OFF}

const
  vtList = -40;
  vtMap  = -41;

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

    constructor CreateI(const Int: IUnknown);
    class function VariantToRef(const Obj: Variant): IUnknown;
    class function ConstToRef(const Obj: TVarRec): IUnknown;
    class function GetTVarRecType(const Obj: Variant): SmallInt;
    class function MakeI(const Int: IUnknown): CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    function GetAsPVariant: PVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    function GetHash: Integer;

    function GetSelfTVarRecType: SmallInt;

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
    constructor CreateDisowned(Obj: TObject); {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
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

    function HasKey(const Key: string): Boolean;

    property AsVariant: Variant read FObj write FObj;
    property Items[const Ind: Variant]: CVariant read GetItems write SetItems; // TODO: Delphi 2006 - default?
    property AsPVariant: PVariant read GetAsPVariant;
    property Hash: Integer read GetHash;
    property IsEmpty: Boolean read GetIsEmpty;
    property IsFull: Boolean read GetIsFull;
    property Size: Integer read GetSize;
    property TVarRecType: SmallInt read GetSelfTVarRecType;
  end;

  CListIterator = {$IFNDEF DELPHI_HAS_RECORDS} object {$ELSE} record {$ENDIF}
  private
    FList, FIterator: IUnknown;
    FPosition: Integer;
    function GetList: CVariant;
  public
    constructor Create(const AList: CVariant);
    destructor Destroy;
    function Next(out Key: Integer; out Value: CVariant): Boolean;
    function NextKey(out Key: Integer): Boolean;
    function NextValue(out Value: CVariant): Boolean;
    property List: CVariant read GetList;
  end;

  CMapIterator = {$IFNDEF DELPHI_HAS_RECORDS} object {$ELSE} record {$ENDIF}
  private
    FMap, FIterator: IUnknown;
    function GetMap: CVariant;
  public
    constructor Create(const AMap: CVariant);
    destructor Destroy;
    function Next(out Key: string; out Value: CVariant): Boolean;
    function NextKey(out Key: string): Boolean;
    function NextValue(out Value: CVariant): Boolean;
    property Map: CVariant read GetMap;
  end;

function CVariantMakeDisowned(Obj: TObject): CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVariantEmpty: CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVariantMake(Obj: TObject): CVariant; overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVariantMake(const Str: string): CVariant;  overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVariantMake(Int: Integer): CVariant; overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVariantMake(Dbl: Double): CVariant;  overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVariantMake(Bol: Boolean): CVariant; overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVariantMakeV(const Vrn: Variant): CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
function CVariantMakeL: CVariant; overload;
function CVariantMakeL(const AItems: array of const): CVariant; overload;
function CVariantMakeM: CVariant; overload;
function CVariantMakeM(const AKeyValues: array of const): CVariant; overload;
function CVariantMakeM(const AKeys, AValues: array of const): CVariant; overload;

implementation

uses
  Collections, SysUtils, Variants, Math;

class function CVariant.VariantToRef(const Obj: Variant): IUnknown;
begin
  Result := nil;
  case TVarData(Obj).VType of
    varEmpty, varNull: Exit;
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
  LIL: IList;
  LIM: IMap;
begin
  Result := vtVariant;
  case TVarData(Obj).VType of
    varEmpty, varNull: Exit;
    varSmallint: Result := vtInteger;
    varInteger: Result := vtInteger;
    varSingle: Result := vtExtended;
    varDouble: Result := vtExtended;
    varCurrency: Result := vtExtended;
    varDate: Result := vtExtended;
    varOleStr: Result := vtWideString;
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
    varString: Result := vtWideString;
    varArray: // TODO: COM array to Collection sequence
  else
    // TODO: custom Variant
    // TODO: Delphi XE2 types
    // TODO: pointer to something
  end;
  if Result = vtInterface then
  begin
    if Supports(LIU, IList, LIL) then
      Result := vtList
    else if Supports(LIU, IMap, LIM) then
      Result := vtMap;
  end;
end;

function CVariant.GetSelfTVarRecType: SmallInt;
begin
  Result := GetTVarRecType(FObj);
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
    vtObject: Result := iown(Obj.VObject);
    vtWideChar: Result := iref(Obj.VWideChar);
    vtPWideChar: Result := iref(Obj.VPWideChar);
    vtAnsiString: Result := iref(AnsiString(Obj.VAnsiString));
    vtCurrency: Result := iref(Obj.VCurrency^);
    vtVariant: Result := VariantToRef(Obj.VVariant^);
    vtInterface: Result := IUnknown(Obj.VInterface);
    vtWideString: Result := iref(WideString(Obj.VWideString));
    vtInt64: Result := iref(Obj.VInt64^);
  else
    // TODO: Delphi XE2 types
  end;
end;

constructor CVariant.CreateI(const Int: IUnknown);
var
  IntObj: IObject;
begin
  if Supports(Int, IObject, IntObj) then
  begin
    case IntObj.delphiType of
    vtInteger: FObj := (IntObj as IInteger).intValue;
    vtExtended: FObj := (IntObj as IDouble).doubleValue;
    vtBoolean: FObj := (IntObj as IBoolean).boolValue;
    vtWideString: FObj := (IntObj as IString).toString;
    else
      FObj := Int;
    end;
  end else
  begin
    FObj := Int;
  end;
end;

class function CVariant.MakeI(const Int: IUnknown): CVariant;
begin
  Result.CreateI(Int);
end;

destructor CVariant.Destroy;
begin
  FObj := Unassigned;
end;

constructor CVariant.CreateDisowned(Obj: TObject);
begin
  FObj := iref(Obj);
end;

constructor CVariant.Create(Obj: TObject);
begin
  FObj := iown(Obj);
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

function CVariantMakeDisowned(Obj: TObject): CVariant;
begin
  Result.FObj := iref(Obj);
end;

function CVariantEmpty: CVariant;
begin
  Result.FObj := Unassigned;
end;

function CVariantMake(Obj: TObject): CVariant;
begin
  Result.FObj := iown(Obj);
end;

function CVariantMake(const Str: string): CVariant;
begin
  Result.FObj := Str;
end;

function CVariantMake(Int: Integer): CVariant;
begin
  Result.FObj := Int;
end;

function CVariantMake(Dbl: Double): CVariant;
begin
  Result.FObj := Dbl;
end;

function CVariantMake(Bol: Boolean): CVariant;
begin
  Result.FObj := Bol;
end;

function CVariantMakeV(const Vrn: Variant): CVariant;
begin
  Result.FObj := Vrn;
end;

function CVariantMakeL: CVariant;
begin
  Result.CreateL;
end;

function CVariantMakeL(const AItems: array of const): CVariant;
begin
  Result.CreateL(AItems);
end;

function CVariantMakeM: CVariant;
begin
  Result.CreateM;
end;

function CVariantMakeM(const AKeyValues: array of const): CVariant;
begin
  Result.CreateM(AKeyValues);
end;

function CVariantMakeM(const AKeys, AValues: array of const): CVariant;
begin
  Result.CreateM(AKeys, AValues);
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
  raise EVariantNotAnArrayError.Create(VarToStr(FObj));
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
  Result := MakeI(GetDeepItem(Indices));
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
  LIU := GetCollection;

  if Supports(LIU, IList, LIL) then
    LIU := LIL.item[Ind]
  else if Supports(LIU, IMap, LIM) then
    LIU := LIM.get(VariantToRef(Ind))
  else
    RaiseNotAnArray;
  LIL := nil; LIM := nil;

  Result := MakeI(LIU);
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
  LIU := GetCollection;

  if Supports(LIU, IMap, LIM) then
    Result := LIM.has(Key)
  else
    RaiseNotAnArray;
  LIM := nil;
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
  end else
    AMap.RaiseNotAnArray;
  LIM := nil;
end;

destructor CMapIterator.Destroy;
begin
  FIterator := nil;
  FMap := nil;
end;

function CMapIterator.Next(out Key: string; out Value: CVariant): Boolean;
var
  Next_Key: string;
begin
  Result := False;
  if not Assigned(FIterator) then Exit;
  if IIterator(FIterator).hasNext then
  begin
    Next_Key := IIterator(FIterator).nextStr;
    Key := Next_Key;
    Value.CreateI(IMap(FMap).Get(Next_Key));
    Result := True;
  end else
    Destroy;
end;

function CMapIterator.NextKey(out Key: string): Boolean;
var
  Next_Key: string;
begin
  Result := False;
  if not Assigned(FIterator) then Exit;
  if IIterator(FIterator).hasNext then
  begin
    Next_Key := IIterator(FIterator).nextStr;
    Key := Next_Key;
    Result := True;
  end else
    Destroy;
end;

function CMapIterator.NextValue(out Value: CVariant): Boolean;
var
  Next_Key: string;
begin
  Result := False;
  if not Assigned(FIterator) then Exit;
  if IIterator(FIterator).hasNext then
  begin
    Next_Key := IIterator(FIterator).nextStr;
    Value.CreateI(IMap(FMap).Get(Next_Key));
    Result := True;
  end else
    Destroy;
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
  end else
    AList.RaiseNotAnArray;
  LIL := nil;
end;

destructor CListIterator.Destroy;
begin
  FList := nil; FIterator := nil; FPosition := 0;
end;

function CListIterator.Next(out Key: Integer; out Value: CVariant): Boolean;
begin
  Result := False;
  if not Assigned(FIterator) then Exit;
  if IIterator(FIterator).hasNext then
  begin
    Key := FPosition;
    Inc(FPosition);
    Value.CreateI(IIterator(FIterator).next);
    Result := True;
  end else
    Destroy;
end;

function CListIterator.NextKey(out Key: Integer): Boolean;
begin
  Result := False;
  if not Assigned(FIterator) then Exit;
  if IIterator(FIterator).hasNext then
  begin
    Key := FPosition;
    Inc(FPosition);
    IIterator(FIterator).next;
    Result := True;
  end else
    Destroy;
end;

function CListIterator.NextValue(out Value: CVariant): Boolean;
begin
  Result := False;
  if not Assigned(FIterator) then Exit;
  if IIterator(FIterator).hasNext then
  begin
    Inc(FPosition);
    Value.CreateI(IIterator(FIterator).next);
    Result := True;
  end else
    Destroy;
end;

function CListIterator.GetList: CVariant;
begin
  Result.CreateI(FList);
end;

end.
