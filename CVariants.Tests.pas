unit CVariants.Tests;

interface
uses
   SysUtils,
   TestFramework,
   CVariants,
   Variants;

{$INCLUDE 'CVariants.DelphiFeatures.inc'}

{$IFNDEF DELPHI_HAS_RECORDS}
{$WARN UNSAFE_TYPE OFF}
{$ENDIF}

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
    procedure TestItemAccess;
    procedure TestUndefinedValues;
    procedure TestListIterator;
    procedure TestMapIterator;
  end;

  TRecursiveTests = class(TTestCase)
  published
    procedure TestOverlay;
    procedure TestDiffEqual;
    procedure TestDiffEqual2;
    procedure TestDiffAdvanced;
    procedure TestEquals;
    procedure TestDiffByEquals;
    procedure TestPatch;
  end;

  TDateTimeTests = class(TTestCase)
  published
    procedure TestBackAndForth;
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
    'List', VList([34, 40, '', 'listitem', nil]),
    'SubMap',
      VMap([
        'er', 0,
        'ui', 1,
        'qx', 7,
        'ow', False
      ])
  ]);
end;

procedure TDeepTests.TestItemAccess;
begin
  CheckEquals(vtMap,      F.VType,                        'F is map');
  CheckEquals(2,          F.Size,                         'F.Size');
  CheckEquals(vtList,     F.Get(['List']).VType,          'F[''List''] is list');
  CheckEquals(5,          F.SizeDeep(['List']),           'F[''List''].Size');
  CheckEquals(vtInteger,  F.Get(['List', 0]).VType,       'F[''List'', 0] is Integer');
  CheckEquals(34,         F.Get(['List', 0]).ToInt,       'F[''List'', 0]');
  CheckEquals(vtInteger,  F.Get(['List', 1]).VType,       'F[''List'', 1] is Integer');
  CheckEquals(40,         F.Get(['List', 1]).ToInt,       'F[''List'', 1]');
  CheckEquals(vtString,   F.Get(['List', 2]).VType,       'F[''List'', 2] is Integer');
  CheckEquals('',         F.Get(['List', 2]).ToString,    'F[''List'', 2]');
  CheckEquals(vtString,   F.Get(['List', 3]).VType,       'F[''List'', 3] is Integer');
  CheckEquals('listitem', F.Get(['List', 3]).ToString,    'F[''List'', 3]');
  CheckEquals(vtEmpty,    F.Get(['List', 4]).VType,       'F[''List'', 4] is Empty');
  CheckEquals(vtMap,      F.Get(['SubMap']).VType,        'F[''SubMap''] is map');
  CheckEquals(4,          F.SizeDeep(['SubMap']),         'F[''SubMap''].Size');
  CheckEquals(vtInteger,  F.VTypeDeep(['SubMap', 'er']),  'F[''er''] is Integer');
  CheckEquals(0,          F.Get(['SubMap', 'er']).ToInt,  'F[''SubMap'', ''er'']');
  CheckEquals(vtInteger,  F.VTypeDeep(['SubMap', 'ui']),  'F[''ui''] is Integer');
  CheckEquals(1,          F.Get(['SubMap', 'ui']).ToInt,  'F[''SubMap'', ''ui'']');
  CheckEquals(vtInteger,  F.VTypeDeep(['SubMap', 'qx']),  'F[''qx''] is Integer');
  CheckEquals(7,          F.Get(['SubMap', 'qx']).ToInt,  'F[''SubMap'', ''qx'']');
  CheckEquals(vtBoolean,  F.VTypeDeep(['SubMap', 'ow']),  'F[''ow''] is Boolean');
  CheckEquals(False,      F.Get(['SubMap', 'ow']).ToBool, 'F[''SubMap'', ''ow'']');
end;

procedure TDeepTests.TestListIterator;
var
  LI: CListIterator;
begin
  LI.Create(F.Get(['List']));
  CheckTrue(LI.Next, 'F[''List'', 0]');
  CheckEquals(0, LI.Key, 'Key');
  CheckEquals(vtInteger, LI.Value.VType, 'F[''List'', 0] is Integer');
  CheckEquals(34, LI.Value.ToInt, 'F[''List'', 0]');
  CheckTrue(LI.Next, 'F[''List'', 1]');
  CheckEquals(1, LI.Key, 'Key');
  CheckEquals(vtInteger, LI.Value.VType, 'F[''List'', 1] is Integer');
  CheckEquals(40, LI.Value.ToInt, 'F[''List'', 1]');
  CheckTrue(LI.Next, 'F[''List'', 2]');
  CheckEquals(2, LI.Key, 'Key');
  CheckEquals(vtString, LI.Value.VType, 'F[''List'', 2] is string');
  CheckEquals('', LI.Value.ToString, 'F[''List'', 2]');
  CheckTrue(LI.Next, 'F[''List'', 3]');
  CheckEquals(3, LI.Key, 'Key');
  CheckEquals(vtString, LI.Value.VType, 'F[''List'', 3] is string');
  CheckEquals('listitem', LI.Value.ToString, 'F[''List'', 3]');
  CheckTrue(LI.Next, 'F[''List'', 4]');
  CheckEquals(4, LI.Key, 'Key');
  CheckEquals(vtEmpty, LI.Value.VType, 'F[''List'', 1] is Empty');
  CheckFalse(LI.Next, 'F[''List'', 5]');
end;

procedure TDeepTests.TestMapIterator;
var
  MI: CMapIterator;
  Count: Integer;
  Occured: array[(er, ui, qx, ow)] of Boolean;
begin
  MI.Create(F.Get(['SubMap']));
  Count := 0;
  Occured[er] := False; Occured[ui] := False;
  Occured[qx] := False; Occured[ow] := False;
  while MI.Next do
  begin
    if MI.Key = 'er' then
    begin
      CheckFalse(Occured[er], 'Occured[er]');
      CheckEquals(vtInteger, MI.Value.VType, 'F[''SubMap'', ''er''] is Integer');
      CheckEquals(0, MI.Value.ToInt, 'F[''SubMap'', ''er'']');
      Occured[er] := True;
    end
    else if MI.Key = 'ui' then
    begin
      CheckFalse(Occured[ui], 'Occured[ui]');
      CheckEquals(vtInteger, MI.Value.VType, 'F[''SubMap'', ''ui''] is Integer');
      CheckEquals(1, MI.Value.ToInt, 'F[''SubMap'', ''ui'']');
      Occured[ui] := True;
    end
    else if MI.Key = 'qx' then
    begin
      CheckFalse(Occured[qx], 'Occured[qx]');
      CheckEquals(vtInteger, MI.Value.VType, 'F[''SubMap'', ''qx''] is Integer');
      CheckEquals(7, MI.Value.ToInt, 'F[''SubMap'', ''qx'']');
      Occured[qx] := True;
    end
    else if MI.Key = 'ow' then
    begin
      CheckFalse(Occured[ow], 'Occured[ow]');
      CheckEquals(vtBoolean, MI.Value.VType, 'F[''SubMap'', ''ow''] is Integer');
      CheckEquals(False, MI.Value.ToBool, 'F[''SubMap'', ''ow'']');
      Occured[ow] := True;
    end;
    Inc(Count);
  end;
  CheckEquals(4, Count, 'F[''SubMap''].Size');
end;

procedure TDeepTests.TestUndefinedValues;
begin
  CheckEquals(vtUndefined, F.VTypeDeep(['Something', 324]), 'F[''Something'', 324]');
  CheckFalse(F.Has(['List', 45]), 'F[''List'', 45]');
  CheckTrue(F.Has(['SubMap', 'ui']), 'F[''SubMap'', ''ui'']');
  CheckEquals(vtUndefined, F.VTypeDeep(['SubMap', 'ui', True]), 'F[''SubMap'', ''ui'', True]');
end;

{ TRecursiveTests }

procedure TRecursiveTests.TestDiffAdvanced;
var
  Obj, Old, Diff: CVariant;
begin
  Obj.CreateM([
    'Servers', VMap([
      'InfoServer', '1.2.3.4',
      'FileServer', '5.6.7.8',
      'VideoServer', '::1'
    ]),
    'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
    'RandomFlag', True,
    'SomeInteger', 3
  ]);
  Old := Obj.Clone;
  Obj.Remove(['Servers', 'FileServer']);
  Obj.Put(['Servers', 'WebServer', '13.14.15.16']);
  Obj.Put(['Servers', 'InfoServer', '9.10.11.12']);
  Obj.Append(['DownloadSources', 'https://ghi.bit/']);
  Obj.Put(['RandomFlag', False]);
  Obj.Put(['SomeNewList'], CList([1, 2, 3, 4, 5]));
  Diff := Obj.DiffFromOld(Old);
  CheckEquals(vtMap, Diff.VType, 'Diff.VType');
  CheckEquals(2, Diff.Size, 'Diff.Size');
  CheckTrue(Diff.HasKey('put'), 'Diff.HasKey(''put'')');
  CheckTrue(Diff.HasKey('overlay'), 'Diff.HasKey(''overlay'')');
  CheckFalse(Diff.HasKey('remove'), 'Diff.HasKey(''remove'')');
  CheckEquals(vtMap, Diff.VTypeDeep(['put']), 'Diff.VTypeDeep([''put''])');
  CheckEquals(1, Diff.SizeDeep(['put']), 'Diff.SizeDeep([''put''])');
  CheckTrue(Diff.Has(['put', 'SomeNewList']), 'Diff.Has([''put'', ''SomeNewList''])');
  CheckEquals(vtList, Diff.VTypeDeep(['put', 'SomeNewList']), 'Diff.VTypeDeep([''put'', ''SomeNewList''])');
  CheckEquals(5, Diff.SizeDeep(['put', 'SomeNewList']), 'Diff.SizeDeep([''put'', ''SomeNewList''])');
  CheckEquals(vtMap, Diff.VTypeDeep(['overlay']), 'Diff.VTypeDeep([''overlay''])');
  CheckEquals(3, Diff.SizeDeep(['overlay']), 'Diff.SizeDeep([''overlay''])');
  CheckTrue(Diff.Has(['overlay', 'Servers']), 'Diff.Has([''overlay'', ''Servers''])');
  CheckTrue(Diff.Has(['overlay', 'DownloadSources']), 'Diff.Has([''overlay'', ''DownloadSources''])');
  CheckTrue(Diff.Has(['overlay', 'RandomFlag']), 'Diff.Has([''overlay'', ''RandomFlag''])');
  CheckEquals(vtMap, Diff.VTypeDeep(['overlay', 'Servers']), 'Diff.VTypeDeep([''overlay'', ''Servers''])');
  CheckEquals(3, Diff.SizeDeep(['overlay', 'Servers']), 'Diff.SizeDeep([''overlay'', ''Servers''])');
  CheckTrue(Diff.Has(['overlay', 'Servers', 'remove']), 'Diff.Has([''overlay'', ''Servers'', ''remove''])');
  CheckTrue(Diff.Has(['overlay', 'Servers', 'put']), 'Diff.Has([''overlay'', ''Servers'', ''put''])');
  CheckTrue(Diff.Has(['overlay', 'Servers', 'overlay']), 'Diff.Has([''overlay'', ''Servers'', ''overlay''])');
  CheckEquals(vtList, Diff.VTypeDeep(['overlay', 'Servers', 'remove']), 'Diff.VTypeDeep([''overlay'', ''Servers'', ''remove''])');
  CheckEquals(1, Diff.SizeDeep(['overlay', 'Servers', 'remove']), 'Diff.SizeDeep([''overlay'', ''Servers'', ''remove''])');
  CheckEquals(vtString, Diff.VTypeDeep(['overlay', 'Servers', 'remove', 0]), 'Diff.VTypeDeep([''overlay'', ''Servers'', ''remove'', 0])');
  CheckEquals('FileServer', Diff.Get(['overlay', 'Servers', 'remove', 0]).ToString, 'Diff.Get([''overlay'', ''Servers'', ''remove'', 0]).ToString');
  CheckEquals(vtMap, Diff.VTypeDeep(['overlay', 'Servers', 'put']), 'Diff.VTypeDeep([''overlay'', ''Servers'', ''put''])');
  CheckEquals(1, Diff.SizeDeep(['overlay', 'Servers', 'put']), 'Diff.SizeDeep([''overlay'', ''Servers'', ''put''])');
  CheckTrue(Diff.Has(['overlay', 'Servers', 'put', 'WebServer']), 'Diff.Has([''overlay'', ''Servers'', ''put'', ''WebServer''])');
  CheckEquals(vtString, Diff.VTypeDeep(['overlay', 'Servers', 'put', 'WebServer']), 'Diff.VTypeDeep([''overlay'', ''Servers'', ''put'', ''WebServer''])');
  CheckEquals('13.14.15.16', Diff.Get(['overlay', 'Servers', 'put', 'WebServer']).ToString, 'Diff.Get([''overlay'', ''Servers'', ''put'', ''WebServer'']).ToString');
  CheckEquals(vtMap, Diff.VTypeDeep(['overlay', 'Servers', 'overlay']), 'Diff.VTypeDeep([''overlay'', ''Servers'', ''overlay''])');
  CheckEquals(1, Diff.SizeDeep(['overlay', 'Servers', 'overlay']), 'Diff.SizeDeep([''overlay'', ''Servers'', ''overlay''])');
  CheckTrue(Diff.Has(['overlay', 'Servers', 'overlay', 'InfoServer']), 'Diff.Has([''overlay'', ''Servers'', ''overlay'', ''InfoServer''])');
  CheckEquals(vtString, Diff.VTypeDeep(['overlay', 'Servers', 'overlay', 'InfoServer']), 'Diff.VTypeDeep([''overlay'', ''Servers'', ''overlay'', ''InfoServer''])');
  CheckEquals('9.10.11.12', Diff.Get(['overlay', 'Servers', 'overlay', 'InfoServer']).ToString, 'Diff.Get([''overlay'', ''Servers'', ''overlay'', ''InfoServer'']).ToString');
  CheckEquals(vtMap, Diff.VTypeDeep(['overlay', 'DownloadSources']), 'Diff.VTypeDeep([''overlay'', ''DownloadSources''])');
  CheckEquals(1, Diff.SizeDeep(['overlay', 'DownloadSources']), 'Diff.SizeDeep([''overlay'', ''DownloadSources''])');
  CheckTrue(Diff.Has(['overlay', 'DownloadSources', 'append']), 'Diff.Has([''overlay'', ''DownloadSources'', ''append''])');
  CheckEquals(vtList, Diff.VTypeDeep(['overlay', 'DownloadSources', 'append']), 'Diff.VTypeDeep([''overlay'', ''DownloadSources'', ''append''])');
  CheckEquals(1, Diff.SizeDeep(['overlay', 'DownloadSources', 'append']), 'Diff.SizeDeep([''overlay'', ''DownloadSources'', ''append''])');
  CheckEquals(vtString, Diff.VTypeDeep(['overlay', 'DownloadSources', 'append', 0]), 'Diff.VTypeDeep([''overlay'', ''DownloadSources'', ''append'', 0])');
  CheckEquals('https://ghi.bit/', Diff.Get(['overlay', 'DownloadSources', 'append', 0]).ToString, 'Diff.Get([''overlay'', ''DownloadSources'', ''append'', 0]).ToString');
  CheckEquals(vtBoolean, Diff.VTypeDeep(['overlay', 'RandomFlag']), 'Diff.VTypeDeep([''overlay'', ''RandomFlag''])');
  CheckEquals(False, Diff.Get(['overlay', 'RandomFlag']).ToBool, 'Diff.Get([''overlay'', ''RandomFlag'']).ToBool');
end;

procedure TRecursiveTests.TestDiffEqual;
var
  Dest, Src: CVariant;
begin
  Src.CreateM([
    'Servers', VMap([
      'InfoServer', '1.2.3.4',
      'FileServer', '5.6.7.8'
    ]),
    'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
    'RandomFlag', True
  ]);
  Dest.CreateM([
    'Servers', VMap([
      'InfoServer', '1.2.3.4',
      'FileServer', '5.6.7.8'
    ]),
    'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
    'RandomFlag', True
  ]);
  CheckEquals(vtEmpty, Dest.DiffFromOld(Src).VType, 'Compare by diff');
end;

procedure TRecursiveTests.TestDiffEqual2;
var
  Dest: CVariant;
begin
  Dest.CreateM([
    'Servers', VMap([
      'InfoServer', '1.2.3.4',
      'FileServer', '5.6.7.8'
    ]),
    'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
    'RandomFlag', True
  ]);
  CheckEquals(vtEmpty, Dest.DiffFromOld(
    CMap([
      'Servers', VMap([
        'InfoServer', '1.2.3.4',
        'FileServer', '5.6.7.8'
      ]),
      'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
      'RandomFlag', True
    ])).VType, 'Compare by diff');
end;

procedure TRecursiveTests.TestEquals;
var
  DummyObj: CVariant;
begin
  CheckTrue(CMap([
    'Servers', VMap([
      'InfoServer', '1.2.3.4',
      'FileServer', '5.6.7.8',
      'VideoServer', '::1'
    ]),
    'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
    'RandomFlag', True,
    'SomeInteger', 3
  ]).Equals(
    CMap([
      'Servers', VMap([
        'InfoServer', '1.2.3.4',
        'FileServer', '5.6.7.8',
        'VideoServer', '::1'
      ]),
      'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
      'RandomFlag', True,
      'SomeInteger', 3
    ])
  ), 'simple equality');

  CheckFalse(CMap([
    'Servers', VMap([
      'InfoServer', '1.2.3.4',
      'FileServer', '5.6.7.8',
      'VideoServer', '::1'
    ]),
    'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
    'RandomFlag', True,
    'SomeInteger', 3
  ]).Equals(
    CMap([
      'Servers', VMap([
        'InfoServer', '1.2.3.4',
        'FileServer', '5.6.7.8',
        'VideoServer', '::2'
      ]),
      'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
      'RandomFlag', True,
      'SomeInteger', 3
    ])
  ), 'simple difference');

  DummyObj.CreateM([
    'Servers', VMap([
      'InfoServer', '1.2.3.4',
      'FileServer', '5.6.7.8',
      'VideoServer', '::1'
    ]),
    'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
    'RandomFlag', True,
    'SomeInteger', 3
  ]);
  DummyObj.Append(['DownloadSources', 'https://def.bit/']);
  DummyObj.Put(['DownloadSources', 1, 'https://abc.bit/']);
  DummyObj.Remove(['DownloadSources', 0]);
  CheckTrue(DummyObj.Equals(
    CMap([
      'Servers', VMap([
        'InfoServer', '1.2.3.4',
        'FileServer', '5.6.7.8',
        'VideoServer', '::1'
      ]),
      'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
      'RandomFlag', True,
      'SomeInteger', 3
    ])
  ), 'different drifts');
  CheckEquals(vtEmpty, DummyObj.DiffFromOld(
    CMap([
      'Servers', VMap([
        'InfoServer', '1.2.3.4',
        'FileServer', '5.6.7.8',
        'VideoServer', '::1'
      ]),
      'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
      'RandomFlag', True,
      'SomeInteger', 3
    ])
  ).VType, 'different drifts; equality by diff');
end;

procedure TRecursiveTests.TestOverlay;
var
  Dest, Src: CVariant;
begin
  Dest.CreateM([
    'Servers', VMap([
      'InfoServer', '1.2.3.4',
      'FileServer', '5.6.7.8'
    ]),
    'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
    'RandomFlag', True
  ]);
  Src.CreateM([
    'Servers', VMap([
      'InfoServer', '9.10.11.12',
      'WebServer', '13.14.15.16'
    ]),
    'DownloadSources', VList(['https://ghi.bit/']),
    'RandomFlag', False
  ]);
  Dest.Overlay(Src);
  CheckEquals(False, Dest.Get(['RandomFlag']).ToBool, 'Dest[''RandomFlag'']');
  CheckEquals(1, Dest.SizeDeep(['DownloadSources']), 'Dest[''DownloadSources''].Size');
  CheckEquals(3, Dest.SizeDeep(['Servers']), 'Dest[''Servers''].Size');
  CheckEquals('9.10.11.12', Dest.Get(['Servers', 'InfoServer']).ToString, 'Dest[''Servers'', ''InfoServer'']');
  CheckEquals('13.14.15.16', Dest.Get(['Servers', 'WebServer']).ToString, 'Dest[''Servers'', ''WebServer'']');
end;

procedure TRecursiveTests.TestDiffByEquals;
begin
  CheckTrue(
    CMap([
      'Servers', VMap([
        'InfoServer', '9.10.11.12',
        'VideoServer', '::1',
        'WebServer', '13.14.15.16'
      ]),
      'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/', 'https://ghi.bit/']),
      'RandomFlag', False,
      'SomeInteger', 3,
      'SomeNewList', VList([1, 2, 3, 4, 5])
    ]).DiffFromOld(
      CMap([
        'Servers', VMap([
          'InfoServer', '1.2.3.4',
          'FileServer', '5.6.7.8',
          'VideoServer', '::1'
        ]),
        'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
        'RandomFlag', True,
        'SomeInteger', 3
      ])
    ).Equals(
      CMap([
        'overlay', VMap([
          'Servers', VMap([
            'remove', VList(['FileServer']),
            'overlay', VMap(['InfoServer', '9.10.11.12']),
            'put', VMap(['WebServer', '13.14.15.16'])
          ]),
          'DownloadSources', VMap([
            'append', VList(['https://ghi.bit/'])
          ]),
          'RandomFlag', False
        ]),
        'put', VMap(['SomeNewList', VList([1, 2, 3, 4, 5])])
      ])
    ), 'diff test by equals');
end;

procedure TRecursiveTests.TestPatch;
var
  O: CVariant;
begin
  O.CreateM([
    'Servers', VMap([
      'InfoServer', '1.2.3.4',
      'FileServer', '5.6.7.8',
      'VideoServer', '::1'
    ]),
    'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/']),
    'RandomFlag', True,
    'SomeInteger', 3
  ]);
  O.ApplyPatch(
    CMap([
      'overlay', VMap([
        'Servers', VMap([
          'remove', VList(['FileServer']),
          'overlay', VMap(['InfoServer', '9.10.11.12']),
          'put', VMap(['WebServer', '13.14.15.16'])
        ]),
        'DownloadSources', VMap([
          'append', VList(['https://ghi.bit/'])
        ]),
        'RandomFlag', False
      ]),
      'put', VMap(['SomeNewList', VList([1, 2, 3, 4, 5])])
    ])
  );
  CheckTrue(
    O.Equals(
      CMap([
        'Servers', VMap([
          'InfoServer', '9.10.11.12',
          'VideoServer', '::1',
          'WebServer', '13.14.15.16'
        ]),
        'DownloadSources', VList(['https://abc.bit/', 'https://def.bit/', 'https://ghi.bit/']),
        'RandomFlag', False,
        'SomeInteger', 3,
        'SomeNewList', VList([1, 2, 3, 4, 5])
      ])
    ), 'patch test by equals');
end;

{ TDateTimeTests }

procedure TDateTimeTests.TestBackAndForth;
var
  Year, Month, Day, Hour, Min, Sec, MSec: Word;
  Obj: CVariant;
begin
  Obj.CreateDT(2005, 05, 20, 18, 22, 09, 071);
  Obj.DecodeDateTime(Year, Month, Day, Hour, Min, Sec, MSec);
  CheckEquals(2005, Year,  'Year');
  CheckEquals(05,   Month, 'Month');
  CheckEquals(20,   Day,   'Day');
  CheckEquals(18,   Hour,  'Hour');
  CheckEquals(22,   Min,   'Min');
  CheckEquals(09,   Sec,   'Sec');
  CheckEquals(071,  MSec,  'MSec');
  CheckEquals(vtDateTime, Obj.VType, 'VType');
  Check(Obj.Equals(CDateTime(2005, 05, 20, 18, 22, 09, 071)));
end;

initialization
  RegisterTests('CVariants',
    [TPrimitiveTests.Suite,
     TListTests.Suite,
     TDeepTests.Suite,
     TRecursiveTests.Suite,
     TDateTimeTests.Suite
    ]);
end.
