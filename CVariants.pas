unit CVariants;

interface

{$INCLUDE 'CVariantDelphiFeatures.inc'}

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
    class function MakeI(const Int: IUnknown): CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    function GetAsPVariant: PVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    function GetHash: Integer;
    function GetItems(const Ind: Variant): CVariant;
    procedure SetItems(const Ind: Variant; const NewObj: CVariant);
  public
    destructor Destroy; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor CreateDisowned(Obj: TObject); {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor Create(Obj: TObject); overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor Create(const Str: string);  overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor Create(Int: Integer); overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor Create(Dbl: Double);  overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor Create(Bol: Boolean); overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    constructor CreateV(const Vrn: Variant); {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    class function MakeDisowned(Obj: TObject): CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    class function Empty: CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    class function Make(Obj: TObject): CVariant; overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    class function Make(const Str: string): CVariant;  overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    class function Make(Int: Integer): CVariant; overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    class function Make(Dbl: Double): CVariant;  overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    class function Make(Bol: Boolean): CVariant; overload; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    class function MakeV(const Vrn: Variant): CVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    function ToString: string;
    function ToInt: Integer;
    function Get(const Bounds: array of const): CVariant; // I failed to make a property out of them
    procedure Put(const Bounds: array of const; const NewObj: CVariant);
    procedure Remove(const Bounds: array of const);
    procedure Insert(const Bounds: array of const; const NewObj: CVariant);
    procedure Append(const Bounds: array of const; const NewObj: CVariant);
    property AsVariant: Variant read FObj write FObj;
    property Items[const Ind: Variant]: CVariant read GetItems write SetItems;
    property AsPVariant: PVariant read GetAsPVariant;
    property Hash: Integer read GetHash;
  end;

implementation

uses
  Collections, SysUtils, Variants;

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

class function CVariant.MakeDisowned(Obj: TObject): CVariant;
begin
  Result.FObj := iref(Obj);
end;

class function CVariant.Empty: CVariant;
begin
  Result.FObj := Unassigned;
end;

class function CVariant.Make(Obj: TObject): CVariant;
begin
  Result.FObj := iown(Obj);
end;

class function CVariant.Make(const Str: string): CVariant;
begin
  Result.FObj := Str;
end;

class function CVariant.Make(Int: Integer): CVariant;
begin
  Result.FObj := Int;
end;

class function CVariant.Make(Dbl: Double): CVariant;
begin
  Result.FObj := Dbl;
end;

class function CVariant.Make(Bol: Boolean): CVariant;
begin
  Result.FObj := Bol;
end;

class function CVariant.MakeV(const Vrn: Variant): CVariant;
begin
  Result.FObj := Vrn;
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

function CVariant.Get(const Bounds: array of const): CVariant;
var
  i: Integer;
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  case TVarData(FObj).VType of
  varUnknown: LIU := IUnknown(TVarData(FObj).VUnknown);
  varDispatch: LIU := IDispatch(TVarData(FObj).VDispatch);
  else
    raise EVariantNotAnArrayError.Create(VarToStr(FObj));
  end;

  for i := Low(Bounds) to High(Bounds) do
  begin
    if Supports(LIU, IObject, LIL) then
    begin
      if Bounds[i].VType <> vtInteger then
        raise EVariantInvalidArgError.Create(stringOf(LIL));
      LIU := LIL.item[Bounds[i].VInteger];
    end else if Supports(LIU, IMap, LIM) then
    begin
      LIU := LIM.get(ConstToRef(Bounds[i]));
    end else
    begin
      raise EInvalidCast.create(VarToStr(FObj));
    end;
    LIL := nil; LIM := nil;
  end;

  Result := MakeI(LIU);
end;

procedure CVariant.Put(const Bounds: array of const; const NewObj: CVariant);
var
  i: Integer;
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  case TVarData(FObj).VType of
  varUnknown: LIU := IUnknown(TVarData(FObj).VUnknown);
  varDispatch: LIU := IDispatch(TVarData(FObj).VDispatch);
  else
    raise EVariantNotAnArrayError.Create(VarToStr(FObj));
  end;

  for i := Low(Bounds) to High(Bounds) - 1 do
  begin
    if Supports(LIU, IObject, LIL) then
    begin
      if Bounds[i].VType <> vtInteger then
        raise EVariantInvalidArgError.Create(stringOf(LIL));
      LIU := LIL.item[Bounds[i].VInteger];
    end else if Supports(LIU, IMap, LIM) then
    begin
      LIU := LIM.get(ConstToRef(Bounds[i]));
    end else
    begin
      raise EInvalidCast.create(VarToStr(FObj));
    end;
    LIL := nil; LIM := nil;
  end;

  if Supports(LIU, IObject, LIL) then
  begin
    if Bounds[High(Bounds)].VType <> vtInteger then
      raise EVariantInvalidArgError.Create(stringOf(LIL));
    LIL.item[Bounds[High(Bounds)].VInteger] := VariantToRef(NewObj.FObj);
  end else if Supports(LIU, IMap, LIM) then
  begin
    LIM.put(ConstToRef(Bounds[High(Bounds)]), VariantToRef(NewObj.FObj));
  end else
  begin
    raise EInvalidCast.create(VarToStr(FObj));
  end;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Remove(const Bounds: array of const);
var
  i: Integer;
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  case TVarData(FObj).VType of
  varUnknown: LIU := IUnknown(TVarData(FObj).VUnknown);
  varDispatch: LIU := IDispatch(TVarData(FObj).VDispatch);
  else
    raise EVariantNotAnArrayError.Create(VarToStr(FObj));
  end;

  for i := Low(Bounds) to High(Bounds) - 1 do
  begin
    if Supports(LIU, IObject, LIL) then
    begin
      if Bounds[i].VType <> vtInteger then
        raise EVariantInvalidArgError.Create(stringOf(LIL));
      LIU := LIL.item[Bounds[i].VInteger];
    end else if Supports(LIU, IMap, LIM) then
    begin
      LIU := LIM.get(ConstToRef(Bounds[i]));
    end else
    begin
      raise EInvalidCast.create(VarToStr(FObj));
    end;
    LIL := nil; LIM := nil;
  end;

  if Supports(LIU, IObject, LIL) then
  begin
    if Bounds[High(Bounds)].VType <> vtInteger then
      raise EVariantInvalidArgError.Create(stringOf(LIL));
    LIL.remove(Bounds[High(Bounds)].VInteger);
  end else if Supports(LIU, IMap, LIM) then
  begin
    LIM.remove(ConstToRef(Bounds[High(Bounds)]));
  end else
  begin
    raise EInvalidCast.create(VarToStr(FObj));
  end;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Insert(const Bounds: array of const; const NewObj: CVariant);
var
  i: Integer;
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  case TVarData(FObj).VType of
  varUnknown: LIU := IUnknown(TVarData(FObj).VUnknown);
  varDispatch: LIU := IDispatch(TVarData(FObj).VDispatch);
  else
    raise EVariantNotAnArrayError.Create(VarToStr(FObj));
  end;

  for i := Low(Bounds) to High(Bounds) - 1 do
  begin
    if Supports(LIU, IObject, LIL) then
    begin
      if Bounds[i].VType <> vtInteger then
        raise EVariantInvalidArgError.Create(stringOf(LIL));
      LIU := LIL.item[Bounds[i].VInteger];
    end else if Supports(LIU, IMap, LIM) then
    begin
      LIU := LIM.get(ConstToRef(Bounds[i]));
    end else
    begin
      raise EInvalidCast.create(VarToStr(FObj));
    end;
    LIL := nil; LIM := nil;
  end;

  if Supports(LIU, IObject, LIL) then
  begin
    if Bounds[High(Bounds)].VType <> vtInteger then
      raise EVariantInvalidArgError.Create(stringOf(LIL));
    LIL.insert(Bounds[High(Bounds)].VInteger, VariantToRef(NewObj.FObj));
  end else if Supports(LIU, IMap, LIM) then
  begin
    LIM.put(ConstToRef(Bounds[High(Bounds)]), VariantToRef(NewObj.FObj));
  end else
  begin
    raise EInvalidCast.create(VarToStr(FObj));
  end;
  LIL := nil; LIM := nil;
end;

procedure CVariant.Append(const Bounds: array of const; const NewObj: CVariant);
var
  i: Integer;
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  case TVarData(FObj).VType of
  varUnknown: LIU := IUnknown(TVarData(FObj).VUnknown);
  varDispatch: LIU := IDispatch(TVarData(FObj).VDispatch);
  else
    raise EVariantNotAnArrayError.Create(VarToStr(FObj));
  end;

  for i := Low(Bounds) to High(Bounds) do
  begin
    if Supports(LIU, IObject, LIL) then
    begin
      if Bounds[i].VType <> vtInteger then
        raise EVariantInvalidArgError.Create(stringOf(LIL));
      LIU := LIL.item[Bounds[i].VInteger];
    end else if Supports(LIU, IMap, LIM) then
    begin
      LIU := LIM.get(ConstToRef(Bounds[i]));
    end else
    begin
      raise EInvalidCast.create(VarToStr(FObj));
    end;
    LIL := nil; LIM := nil;
  end;

  if Supports(LIU, IObject, LIL) then
  begin
    LIL.add(VariantToRef(NewObj.FObj));
  end else
  begin
    raise EInvalidCast.create(VarToStr(FObj));
  end;
  LIL := nil; LIM := nil;
end;

function CVariant.GetItems(const Ind: Variant): CVariant;
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  case TVarData(FObj).VType of
  varUnknown: LIU := IUnknown(TVarData(FObj).VUnknown);
  varDispatch: LIU := IDispatch(TVarData(FObj).VDispatch);
  else
    raise EVariantNotAnArrayError.Create(VarToStr(FObj));
  end;

  if Supports(LIU, IObject, LIL) then
  begin
    LIU := LIL.item[Ind];
  end else if Supports(LIU, IMap, LIM) then
  begin
    LIU := LIM.get(VariantToRef(Ind));
  end else
  begin
    raise EInvalidCast.create(VarToStr(FObj));
  end;
  LIL := nil; LIM := nil;

  Result := MakeI(LIU);
end;

procedure CVariant.SetItems(const Ind: Variant; const NewObj: CVariant);
var
  LIU: IUnknown;
  LIL: IList;
  LIM: IMap;
begin
  case TVarData(FObj).VType of
  varUnknown: LIU := IUnknown(TVarData(FObj).VUnknown);
  varDispatch: LIU := IDispatch(TVarData(FObj).VDispatch);
  else
    raise EVariantNotAnArrayError.Create(VarToStr(FObj));
  end;

  if Supports(LIU, IObject, LIL) then
  begin
    LIL.item[Ind] := VariantToRef(NewObj.FObj);
  end else if Supports(LIU, IMap, LIM) then
  begin
    LIM.put(VariantToRef(Ind), VariantToRef(NewObj.FObj));
  end else
  begin
    raise EInvalidCast.create(VarToStr(FObj));
  end;
  LIL := nil; LIM := nil;
end;

end.
