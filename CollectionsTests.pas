{ $Id: CollectionsTests.pas,v 1.6 2001/12/17 04:40:52 juanco Exp $ }
{: DUnit: An XTreme testing framework for Delphi programs.
   @author  The DUnit Group.
   @version $Revision: 1.6 $
}
(*
 * The contents of this file are subject to the Mozilla Public
 * License Version 1.1 (the "License"); you may not use this file
 * except in compliance with the License. You may obtain a copy of
 * the License at http://www.mozilla.org/MPL/
 *
 * Software distributed under the License is distributed on an "AS
 * IS" basis, WITHOUT WARRANTY OF ANY KIND, either express or
 * implied. See the License for the specific language governing
 * rights and limitations under the License.
 *
 * The Original Code is DUnit.
 *
 * The Initial Developers of the Original Code are Kent Beck, Erich Gamma,
 * and Juancarlo Añez.
 * Portions created The Initial Developers are Copyright (C) 1999-2000.
 * Portions created by The DUnit Group are Copyright (C) 2000.
 * All rights reserved.
 *
 * Contributor(s):
 * Kent Beck <kentbeck@csi.com>
 * Erich Gamma <Erich_Gamma@oti.com>
 * Juanco Añez <juanco@users.sourceforge.net>
 * Chris Morris <chrismo@users.sourceforge.net>
 * Jeff Moore <JeffMoore@users.sourceforge.net>
 * The DUnit group at SourceForge <http://dunit.sourceforge.net>
 *
 *)
unit CollectionsTests;
interface
uses
   SysUtils,
   TestFramework,
   Collections;

type
  TListTests = class(TTestCase)
  protected
    procedure doEmptyTests(list :IList);
    procedure doSingleItemTests(list :IList);
    procedure testIndexOf(list :IList);
  published
    // methods are published so RTTI is generated
    // methods are virtual so the compiler doesn't eliminate them
    procedure testEmptyArrayList;   virtual;
    procedure testSingleItemArrayList;   virtual;
    procedure testEmptyLinkedList;  virtual;
    procedure testSingleItemLinkedList;  virtual;
    procedure testArrayListIndexOf;
  end;

  TMapTests = class(TTestCase)
  protected
    function  buildTreeMap :ISortedMap;
    procedure testMap(map :IMap);
  published
    // methods are published so RTTI is generated
    // methods are virtual so the compiler doesn't eliminate them
    procedure testHashMap;          virtual;
    procedure testTreeMap;          virtual;
    procedure testSortedMap;        virtual;
    procedure testSortedMapValues;  virtual;
  end;

  TSetTests = class(TTestCase)
  protected
    function  buildTreeSet :ISortedSet;
    procedure testSet(aset :ISet);
  published
    // methods are published so RTTI is generated
    // methods are virtual so the compiler doesn't eliminate them
    procedure testHashSet;          virtual;
    procedure testTreeSet;          virtual;
    procedure testSortedSet;        virtual;
    procedure testSortedSetValues;  virtual;
  end;

  THeapTests = class(TTestCase)
    heap :IStack;
    procedure setUp;    override;
    procedure tearDown; override;
  published
    procedure desceningOrder;    virtual;
    procedure randomOrder;       virtual;
  end;


implementation

{ TListTests }

procedure TListTests.doEmptyTests(list: IList);
var
   i :IIterator;
begin
  check(list.isEmpty,                       'isEmpty');
  checkEquals(0, list.size,                 'size');
  check(not list.iterator.hasNext,          'iterator at end');

  list.add('anItem');
  checkEquals(1, list.size,                 'size');
  i := list.iterator;
  check(i.hasNext,                          'iterator valid');
  i.next;
  check(not i.hasNext,                      'iterator only one item');
  checkEquals(stringOf(list.iterator.next), 'anItem',
                                            'iterator only one item');

  list.remove(0);
  check(list.isEmpty,                       'isEmpty after remove');
  checkEquals(0, list.size,                 'size after remove');
  check(not list.iterator.hasNext,          'iterator at end after remove');
end;

procedure TListTests.doSingleItemTests(list: IList);
begin
  doEmptyTests(list);
  list.clear;
  check(list.isEmpty,                       'isEmpty');
  checkEquals(0, list.size,                 'size');

  list.add('an item');
  checkEquals(1, list.size,                 'size');
  checkEquals('an item', stringOf(list.at(0)));

  list.remove(0);
  check(list.isEmpty,                       'isEmpty');
  checkEquals(0, list.size,                 'size');

  list.add('item a');
  list.add('item b');
  checkEquals(2, list.size,                 'size');
  checkEquals('item a', stringOf(list.at(0)));

  list.remove(0);
  checkEquals(1, list.size,                 'size');
  checkEquals('item b', stringOf(list.at(0)));

  list.insert(0, 'item c');
  checkEquals(2, list.size,                 'size');
  checkEquals('item c', stringOf(list.at(0)));

  list.remove(1);
  checkEquals(1, list.size,                 'size');
  checkEquals('item c', stringOf(list.at(0)));
end;

procedure TListTests.testArrayListIndexOf;
begin
  testIndexOf(TArrayList.create);
end;

procedure TListTests.testEmptyArrayList;
begin
  doEmptyTests(TArrayList.create);
end;

procedure TListTests.testEmptyLinkedList;
begin
  doEmptyTests(TLinkedList.Create);
end;

procedure TListTests.testIndexOf(list: IList);
var
  i    :integer;
begin
  for i := 0 to 99 do
    list.add(iref(i));
  checkEquals(100, list.size);
  checkEquals(50, list.indexOf(iref(50)));
end;

procedure TListTests.testSingleItemArrayList;
begin
  doSingleItemTests(TArrayList.create);
end;

procedure TListTests.testSingleItemLinkedList;
begin
  doSingleItemTests(TLinkedList.create);
end;

{ TMapTests }

function TMapTests.buildTreeMap: ISortedMap;
begin
  result := TTreeMap.create;
  result.put('z','z');
  result.put('c','c');
  result.put('x','x');
  result.put('a','a');
  result.put('m','m');
end;

procedure TMapTests.testMap(map: IMap);
var
  a, b, c :IUnknown;
  o1, o2  :TObject;
  x       :IUnknown;
begin
  a := iref('a');
  b := iref('b');
  c := iref('c');

  checkEquals(map.size, 0);

  map.put(a, a);
  checkSame(a, map.get(a));
  checkEquals(map.size, 1);

  map.put(b, b);
  checkSame(a, map.get(a));
  checkSame(b, map.get(b));
  checkEquals(map.size, 2);

  map.put(c, c);
  checkSame(a, map.get(a));
  checkSame(b, map.get(b));
  checkSame(c, map.get(c));
  checkEquals(map.size, 3);

  map.remove(b);
  checkSame(a, map.get(a));
  checkNull(map.get(b));
  checkSame(c, map.get(c));
  checkEquals(map.size, 2);

  map.remove(a);
  checkNull(map.get(a));
  checkNull(map.get(b));
  checkSame(c, map.get(c));
  checkEquals(map.size, 1);

  map.remove(c);
  checkNull(map.get(a));
  checkNull(map.get(b));
  checkNull(map.get(c));
  checkEquals(map.size, 0);

  o1 := TObject.create;
  o2 := TObject.create;
  map.put(iref(o1), iref(o2));
  x := map.get(iref(o1));
  check(Collections.equal(x, iref(o2)));
  checkNull(map.get(iref(o2)));

  map.put(iref(o2), self);
  check(Collections.equal(self, map.get(iref(o2))));
  map.remove(iref(o2));
  map.remove(iref(o1));
  checkEquals(0, map.size);

  map.put(iref(101), x);
  check(Collections.equal(map.get(iref(101)), x));
  check(not Collections.equal(map.get(iref(101)), a));
end;

procedure TMapTests.testHashMap;
begin
   testMap(THashMap.create)
end;

procedure TMapTests.testSortedMap;
var
  s :ISortedMap;
  i :IIterator;
  o :IUnknown;
begin
  s := buildTreeMap;

  i := s.keys;
  o := i.next;
  check(Collections.equal(iref('a'), o),                'check key');
  check(Collections.equal(iref('a'), s.get(o)),         'check value');
  check(Collections.equal(iref('a'), s.get(iref('a'))), 'check value 2');
end;

procedure TMapTests.testSortedMapValues;
var
  s :IMap;
  i :IIterator;
begin
  s := buildTreeMap;
  i := s.values.iterator;
  check(Collections.equal(iref('a'), i.next),         'a');
  check(Collections.equal(iref('c'), i.next),         'c');
  check(Collections.equal(iref('m'), i.next),         'm');
  check(Collections.equal(iref('x'), i.next),         'x');
  check(Collections.equal(iref('z'), i.next),         'z');
  check(not i.hasNext,                   'eol');
end;



procedure TMapTests.testTreeMap;
begin
    testMap(TTreeMap.create)
end;

{ TSetTests }

function TSetTests.buildTreeSet: ISortedSet;
begin
  result := TTreeSet.create;
  result.add('z');
  result.add('c');
  result.add('x');
  result.add('a');
  result.add('m');
end;

procedure TSetTests.testHashSet;
begin
  testSet(THashSet.create)
end;

procedure TSetTests.testSet(aset: ISet);
var
  a, b, c :IUnknown;
begin
  a := iref('a');
  b := iref('b');
  c := iref('c');

  checkEquals(aset.size, 0);
  check(aset.isEmpty);

  aset.add(b);
  check(not aset.has(a), 'not has a');
  check(aset.has(b),     'has b');
  check(not aset.has(c), 'not has c');
  checkEquals(aset.size, 1);

  aset.add(a);
  check(aset.has(a), 'has a');
  check(aset.has(b), 'has b');
  check(not aset.has(c), 'not has c');
  checkEquals(aset.size, 2);

  aset.add(c);
  check(aset.has(a), 'has a');
  check(aset.has(b), 'has b');
  check(aset.has(c), 'has c');
  checkEquals(aset.size, 3);

  aset.remove(b);
  check(aset.has(a), 'has a');
  check(not aset.has(b), 'not has b');
  check(aset.has(c), 'has c');
  checkEquals(aset.size, 2);

  aset.remove(a);
  check(not aset.has(a), 'not has a');
  check(not aset.has(b), 'not has b');
  check(aset.has(c), 'has c');
  checkEquals(aset.size, 1);

  aset.remove(c);
  check(not aset.has(a), 'not has a');
  check(not aset.has(b), 'not has b');
  check(not aset.has(c), 'not has c');
  checkEquals(aset.size, 0);
  check(aset.isEmpty);
end;

procedure TSetTests.testSortedSet;
var
  s :ISortedSet;
  i :IIterator;
begin
  s := buildTreeSet;

  check(Collections.equal(iref('a'), s.first), 'check first');
  check(Collections.equal(iref('z'), s.last), 'check last');

  i := s.iterator;
  check(Collections.equal(iref('a'), i.next), 'check a');
  check(Collections.equal(iref('c'), i.next), 'check c');
  check(Collections.equal(iref('m'), i.next), 'check m');
  check(Collections.equal(iref('x'), i.next), 'check x');
  check(Collections.equal(iref('z'), i.next), 'check z');
  check(not i.hasNext, 'eos');
end;

procedure TSetTests.testSortedSetValues;
var
  s :ISet;
  i :IIterator;
begin
  s := buildTreeSet;
  i := s.iterator;
  check(Collections.equal(iref('a'), i.next),         'a');
  check(Collections.equal(iref('c'), i.next),         'c');
  check(Collections.equal(iref('m'), i.next),         'm');
  check(Collections.equal(iref('x'), i.next),         'x');
  check(Collections.equal(iref('z'), i.next),         'z');
  check(not i.hasNext,                   'eol');
end;

procedure TSetTests.testTreeSet;
begin
  testSet(TTreeSet.create)
end;

{ THeapTests }

procedure THeapTests.setUp;
begin
  heap := THeap.create(TArrayList.create);
  System.RandSeed := 0;
end;

procedure THeapTests.tearDown;
begin
  heap := nil;
end;
procedure THeapTests.desceningOrder;
var
  i    :integer;
  last :integer;
begin
  for i := 200 downto 1 do
    heap.push(iref(i*100));

  checkNotNull(heap.remove(iref(50*100)), 'remove does not find items');
  checkNull(heap.remove(iref(5000*100)), 'remove found unavailable item');

  last := 0;
  while not heap.isEmpty do
  begin
    check(last <= (heap.pop as IInteger).intValue);
  end;
end;

procedure THeapTests.randomOrder;
var
  i    :integer;
  n    :integer;
  last :integer;
begin
  for i := 1 to 200 do
    heap.push(iref(random(10000)));

  checkNotNull(heap.remove(heap.items[100]), 'remove does not find items');
  checkNull(heap.remove(iref(5000*100)), 'remove found unavailable item');

  last := 0; i := 0;
  while not heap.isEmpty do
  begin
    inc(i);
    n := (heap.pop as IInteger).intValue;
    check(last <= n, format('%d: %d > %d', [i, last, n]));
    check(THeap((heap as IDelphiObject).obj).check, intToStr(i)); 
  end;
end;

initialization
  RegisterTests('',       [TListTests.Suite,
                          TMapTests.Suite,
                          TSetTests.Suite,
                          THeapTests.Suite
                          ]);

end.
