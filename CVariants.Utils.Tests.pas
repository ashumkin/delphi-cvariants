unit CVariants.Utils.Tests;

interface
uses
   SysUtils,
   TestFramework,
   CVariants,
   CVariants.Utils,
   Variants,
   Classes,
   WideStrings;

{$INCLUDE 'CVariants.DelphiFeatures.inc'}

{$IFNDEF DELPHI_HAS_RECORDS}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

type
  TStringListTests = class(TTestCase)
  protected
    FTStrings: TStringList;
    FTWideStrings: TWideStringList;
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure TestFromTStrings;
    procedure TestToTStrings;
    procedure TestMapFromTStrings;
    procedure TestMapToTStrings;
  end;

implementation

{ TStringListTests }

procedure TStringListTests.SetUp;
begin
  inherited;
  FTStrings := TStringList.Create;
  FTWideStrings := TWideStringList.Create;
end;

procedure TStringListTests.TearDown;
begin
  inherited;
  FreeAndNil(FTStrings);
  FreeAndNil(FTWideStrings);
end;

procedure TStringListTests.TestFromTStrings;
begin
  FTStrings.Add('dfgsdfgsdg');
  FTStrings.Add('bxbrwerhsdf');
  Check(CListFromStrings(FTStrings).Equals(CList([
    'dfgsdfgsdg', 'bxbrwerhsdf'
  ])), 'TStrings; compared by .Equals()');
  FTWideStrings.Add('dfgsdfgsdg');
  FTWideStrings.Add('bxbrwerhsdf');
  Check(CListFromWideStrings(FTWideStrings).Equals(CList([
    'dfgsdfgsdg', 'bxbrwerhsdf'
  ])), 'TWideStrings; compared by .Equals()');
end;

procedure TStringListTests.TestMapFromTStrings;
begin
  FTStrings.Add('dfgsdfgsdg=dbsdfb');
  FTStrings.Add('bxbrwerhsdf=34twgwwerg');
  Check(CMapFromStrings(FTStrings).Equals(CMap([
    'dfgsdfgsdg', 'dbsdfb',
    'bxbrwerhsdf', '34twgwwerg'
  ])), 'TStrings; compared by .Equals()');
  FTWideStrings.Add('dfgsdfgsdg=dbsdfb');
  FTWideStrings.Add('bxbrwerhsdf=34twgwwerg');
  Check(CMapFromWideStrings(FTWideStrings).Equals(CMap([
    'dfgsdfgsdg', 'dbsdfb',
    'bxbrwerhsdf', '34twgwwerg'
  ])), 'TWideStrings; compared by .Equals()');
end;

procedure TStringListTests.TestMapToTStrings;
begin
  CMapToStrings(CMap([
    'dfgsdfgsdg', 'dbsdfb',
    'bxbrwerhsdf', '34twgwwerg'
  ]), FTStrings);
  CheckEquals(2, FTStrings.Count);
  CheckEquals('dbsdfb', FTStrings.Values['dfgsdfgsdg']);
  CheckEquals('34twgwwerg', FTStrings.Values['bxbrwerhsdf']);
  CMapToWideStrings(CMap([
    'dfgsdfgsdg', 'dbsdfb',
    'bxbrwerhsdf', '34twgwwerg'
  ]), FTWideStrings);
  CheckEquals(2, FTWideStrings.Count);
  CheckEquals('dbsdfb', FTWideStrings.Values['dfgsdfgsdg']);
  CheckEquals('34twgwwerg', FTWideStrings.Values['bxbrwerhsdf']);
end;

procedure TStringListTests.TestToTStrings;
begin
  CListToStrings(CList([
    'dfgsdfgsdg', 'bxbrwerhsdf'
  ]), FTStrings);
  CheckEquals(2, FTStrings.Count, 'FTStrings.Count');
  CheckEquals('dfgsdfgsdg', FTStrings[0], 'FTStrings[0]');
  CheckEquals('bxbrwerhsdf', FTStrings[1], 'FTStrings[1]');
  CListToWideStrings(CList([
    'dfgsdfgsdg', 'bxbrwerhsdf'
  ]), FTWideStrings);
  CheckEquals(2, FTWideStrings.Count, 'FTWideStrings.Count');
  CheckEquals('dfgsdfgsdg', FTWideStrings[0], 'FTWideStrings[0]');
  CheckEquals('bxbrwerhsdf', FTWideStrings[1], 'FTWideStrings[1]');
end;

initialization
  RegisterTests('CVariants.Utils',
    [TStringListTests.Suite
    ]);
end.
