unit CVariantTests;

interface
uses
   SysUtils,
   TestFramework,
   CVariants;

{$WARN UNSAFE_TYPE OFF}

type
  TPrimitiveTests = class(TTestCase)
  published
    procedure TestInteger;
    procedure TestString;
  end;

  TListTests = class(TTestCase)
  protected
    FList: CVariant;
    procedure SetUp; override;
  published
    procedure TestSize;
    procedure TestItemAccess;
  end;

  TMapTests = class(TTestCase)
  protected
    FMap: CVariant;
    procedure SetUp; override;
  published
    procedure TestItemAccess;
    procedure TestHas;
  end;

  TDeepTests = class(TTestCase)
  protected
    F: CVariant;
    procedure SetUp; override;
  published
    procedure Test1;
  end;

implementation


{ TPrimitiveTests }

procedure TPrimitiveTests.TestInteger;
var
  I1, I2: CVariant;
begin
  I1.Create(3);
  CheckEquals(3, I1.ToInt, 'Create, ToInt');
  I2 := I1;
  CheckEquals(3, I2.ToInt, 'Create, ToInt, Assign');
  CheckEquals(3, CVar(3).ToInt, 'Make');
  I1.Destroy;
  I2.Destroy;
end;

procedure TPrimitiveTests.TestString;
var
  I1, I2: CVariant;
begin
  I1.Create('cbd');
  CheckEquals('cbd', I1.ToString, 'Create, ToString');
  I2 := I1;
  CheckEquals('cbd', I2.ToString, 'Create, ToString, Assign');
  CheckEquals('cbd', CVar('cbd').ToString, 'Make');
  I1.Destroy;
  I2.Destroy;
end;

{ TListTests }

procedure TListTests.SetUp;
begin
  inherited;
  FList.CreateL(['123', 17, True]);
end;

procedure TListTests.TestItemAccess;
begin
  CheckEquals('123', FList.Get([0]).ToString, '[0].ToString');
  CheckEquals(17, FList.Get([1]).ToInt, '[1].ToInt');
  CheckEquals(True, FList.Get([2]).ToBool, '[2].ToBool');
end;

procedure TListTests.TestSize;
begin
  CheckEquals(3, FList.Size, 'List.Size');
end;

{ TMapTests }

procedure TMapTests.SetUp;
begin
  inherited;
  FMap.CreateM(['abc', True, 'qore', 4]);
end;

procedure TMapTests.TestHas;
begin
  CheckTrue(FMap.HasKey('abc'), 'Has key ''abc''');
  CheckTrue(FMap.HasKey('qore'), 'Has key ''qore''');
  CheckFalse(FMap.HasKey('12312'), 'Has key: ''12312''');
end;

procedure TMapTests.TestItemAccess;
begin
  CheckEquals(True, FMap.Get(['abc']).ToBool, '[''abc''].ToBool');
  CheckEquals(4, FMap.Get(['qore']).ToInt, '[''qore''].ToInt');
end;

{ TDeepTests }

procedure TDeepTests.SetUp;
begin
  inherited;
  F.CreateM([
    'List', VList([34, 40, '', 'listitem']),
    'SubMap',
      VMap([
        'er', 0,
        'ui', 1,
        'qx', 7,
        'ow', False
      ])
  ]);
end;

procedure TDeepTests.Test1;
begin
  CheckEquals(vtMap,        F.VType,                        'F is map');
  CheckEquals(2,            F.Size,                         'F.Size');
  CheckEquals(vtList,       F.Get(['List']).VType,          'F[''List''] is list');
  CheckEquals(4,            F.SizeDeep(['List']),           'F[''List''].Size');
  CheckEquals(vtInteger,    F.Get(['List', 0]).VType,       'F[''List'', 0] is Integer');
  CheckEquals(34,           F.Get(['List', 0]).ToInt,       'F[''List'', 0]');
  CheckEquals(vtInteger,    F.Get(['List', 1]).VType,       'F[''List'', 1] is Integer');
  CheckEquals(40,           F.Get(['List', 1]).ToInt,       'F[''List'', 1]');
  CheckEquals(vtWideString, F.Get(['List', 2]).VType,       'F[''List'', 2] is Integer');
  CheckEquals('',           F.Get(['List', 2]).ToString,    'F[''List'', 2]');
  CheckEquals(vtWideString, F.Get(['List', 3]).VType,       'F[''List'', 3] is Integer');
  CheckEquals('listitem',   F.Get(['List', 3]).ToString,    'F[''List'', 3]');
  CheckEquals(vtMap,        F.Get(['SubMap']).VType,        'F[''SubMap''] is map');
  CheckEquals(4,            F.SizeDeep(['SubMap']),         'F[''SubMap''].Size');
  CheckEquals(vtInteger,    F.VTypeDeep(['SubMap', 'er']),  'F[''er''] is Integer');
  CheckEquals(0,            F.Get(['SubMap', 'er']).ToInt,  'F[''SubMap'', ''er'']');
  CheckEquals(vtInteger,    F.VTypeDeep(['SubMap', 'ui']),  'F[''ui''] is Integer');
  CheckEquals(1,            F.Get(['SubMap', 'ui']).ToInt,  'F[''SubMap'', ''ui'']');
  CheckEquals(vtInteger,    F.VTypeDeep(['SubMap', 'qx']),  'F[''qx''] is Integer');
  CheckEquals(7,            F.Get(['SubMap', 'qx']).ToInt,  'F[''SubMap'', ''qx'']');
  CheckEquals(vtBoolean,    F.VTypeDeep(['SubMap', 'ow']),  'F[''ow''] is Boolean');
  CheckEquals(False,        F.Get(['SubMap', 'ow']).ToBool, 'F[''SubMap'', ''ow'']');
end;

initialization
  RegisterTests('CVariants',
    [TPrimitiveTests.Suite,
     TListTests.Suite,
     TDeepTests.Suite
    ]);
end.
