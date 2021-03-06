unit CVariants.Utils;

interface
uses
  SysUtils,
  CVariants,
  Variants,
  Classes,
  WideStrings;

{$INCLUDE 'CVariants.DelphiFeatures.inc'}

{$IFNDEF DELPHI_HAS_RECORDS}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

// conversion

function CListFromStrings(AList: TStrings): CVariant;
procedure CListToStrings(const AList: CVariant; AOutput: TStrings);

function CMapFromStrings(AMap: TStrings): CVariant;
procedure CMapToStrings(const AMap: CVariant; AOutput: TStrings);

function CListFromWideStrings(AList: TWideStrings): CVariant;
procedure CListToWideStrings(const AList: CVariant; AOutput: TWideStrings);

function CMapFromWideStrings(AMap: TWideStrings): CVariant;
procedure CMapToWideStrings(const AMap: CVariant; AOutput: TWideStrings);

implementation

function CListFromStrings(AList: TStrings): CVariant;
var
  i, L: Integer;
begin
  Result.CreateL;
  if not Assigned(AList) then Exit;
  L := AList.Count;
  for i := 0 to L - 1 do
    Result.Append(CVar(AList[i]));
end;

procedure CListToStrings(const AList: CVariant; AOutput: TStrings);
var
  LI: CListIterator;
begin
  if not Assigned(AOutput) then
    raise EVariantInvalidArgError.Create('Output TStrings is nil');

  LI.Create(AList);

  AOutput.Clear;
  AOutput.Capacity := AList.Size;

  while LI.Next do
    AOutput.Add(LI.Value.ToString);
end;

function CMapFromStrings(AMap: TStrings): CVariant;
var
  i, L: Integer;
begin
  Result.CreateM;
  if not Assigned(AMap) then Exit;
  L := AMap.Count;
  for i := 0 to L - 1 do
    Result.Put([AMap.Names[i]], CVar(AMap.ValueFromIndex[i]));
end;

procedure CMapToStrings(const AMap: CVariant; AOutput: TStrings);
var
  MI: CMapIterator;
begin
  if not Assigned(AOutput) then
    raise EVariantInvalidArgError.Create('Output TStrings is nil');

  MI.Create(AMap);

  AOutput.Clear;

  while MI.Next do
    AOutput.Add(MI.Key + AOutput.NameValueSeparator + MI.Value.ToString);
end;

function CListFromWideStrings(AList: TWideStrings): CVariant;
var
  i, L: Integer;
begin
  Result.CreateL;
  if not Assigned(AList) then Exit;
  L := AList.Count;
  for i := 0 to L - 1 do
    Result.Append(CVar(AList[i]));
end;

procedure CListToWideStrings(const AList: CVariant; AOutput: TWideStrings);
var
  LI: CListIterator;
begin
  if not Assigned(AOutput) then
    raise EVariantInvalidArgError.Create('Output TWideStrings is nil');

  LI.Create(AList);

  AOutput.Clear;
  AOutput.Capacity := AList.Size;

  while LI.Next do
    AOutput.Add(LI.Value.ToString);
end;

function CMapFromWideStrings(AMap: TWideStrings): CVariant;
var
  i, L: Integer;
begin
  Result.CreateM;
  if not Assigned(AMap) then Exit;
  L := AMap.Count;
  for i := 0 to L - 1 do
    Result.Put([AMap.Names[i]], CVar(AMap.ValueFromIndex[i]));
end;

procedure CMapToWideStrings(const AMap: CVariant; AOutput: TWideStrings);
var
  MI: CMapIterator;
begin
  if not Assigned(AOutput) then
    raise EVariantInvalidArgError.Create('Output TWideStrings is nil');

  MI.Create(AMap);

  AOutput.Clear;

  while MI.Next do
    AOutput.Add(MI.Key + AOutput.NameValueSeparator + MI.Value.ToString);
end;

end.
