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


    class function VariantToRef(const Obj: Variant): IUnknown;
    class function ConstToRef(const Obj: TVarRec): IUnknown;
    function GetAsPVariant: PVariant; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
    function GetHash: Integer;
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
    property AsVariant: Variant read FObj write FObj;
    property AsPVariant: PVariant read GetAsPVariant;
    property Hash: Integer read GetHash;
  end;

implementation

uses
  Collections, Variants;

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

end.
