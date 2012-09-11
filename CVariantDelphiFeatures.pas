unit
  CVariantDelphiFeatures;

interface

{$INCLUDE 'CVariantDelphiFeatures.inc'}

{$IFNDEF DELPHI_IS_UNICODE}
type
  UnicodeString = type WideString;

function UTF8ToUnicodeString(const S: UTF8String): UnicodeString; {$IFDEF DELPHI_HAS_INLINE} inline; {$ENDIF}
// Decode is deprecated, Encode is not
// because Encode can decide on its argument type

{$ENDIF}


{$IFNDEF DELPHI_HAS_UINT64}
type
  UInt64 = type Int64;
{$ENDIF}

implementation

{$IFNDEF DELPHI_IS_UNICODE}
function UTF8ToUnicodeString(const S: UTF8String): UnicodeString;
begin
  Result := UTF8Decode(S);
end;
{$ENDIF}


end.
