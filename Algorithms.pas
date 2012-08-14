{ $Id: Algorithms.pas,v 1.2 2000/08/01 17:21:11 juanco Exp $ }
{: Collections: A Delphi port of the Java Collections library.
   @author  Juancarlo Añez.
   @version $Revision: 1.2 $
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
 * The Original Code is Collections.
 *
 * The Initial Developer of the Original Code is Juancarlo Añez.
 * Portions created The Initial Developers are Copyright (C) 1999-2000.
 * Portions created by The DUnit Group are Copyright (C) 2000.
 * All rights reserved.
 *
 * Contributor(s):
 * Juanco Añez <juanco@users.sourceforge.net>
 * Chris Morris <chrismo@users.sourceforge.net>
 * The DUnit group at SourceForge <http://dunit.sourceforge.net>
 *
 *)
unit Algorithms;
interface
uses
   Collections;

const
  rcs_id :string = '@(#)$Id: Algorithms.pas,v 1.2 2000/08/01 17:21:11 juanco Exp $';

type
   recurse_t = (rt_donotrecurse, rt_recurse);

   IUnaryFunction = interface
   ['{BBFBE12C-F88F-4C6A-946E-0E049FFC712E}']
      function run( obj: IUnknown ): IUnknown;
   end;

   IBinaryFunction = interface
   ['{9082D051-94F0-4008-86B2-F9DC92EA01CA}']
      function run( obj1, obj2: IUnknown ): IUnknown;
   end;

function forEach(coll :ICollection; func :IUnaryFunction) :IUnaryFunction; overload;
function inject( coll :ICollection; obj: IUnknown; func: IBinaryFunction ): IUnknown; overload;

function count( coll :ICollection; recurse :recurse_t = rt_donotrecurse) :Integer; overload;
function count( map  :IMap; recurse :recurse_t = rt_donotrecurse) :Integer; overload;


implementation

function forEach(coll :ICollection; func :IUnaryFunction) :IUnaryFunction; overload;
var
  i :IIterator;
begin
   i := coll.iterator;
   while i.hasNext do
       func.run(i.next);
   result := func
end;

function inject( coll :ICollection; obj: IUnknown; func: IBinaryFunction ): IUnknown; overload;
var
  i :IIterator;
begin
   i := coll.iterator;
   while i.hasNext do
       obj := func.run(obj, i.next);
   result := obj
end;

function count( coll :ICollection; recurse :recurse_t = rt_donotrecurse) :Integer;
var
  i   :IIterator;
  obj :IUnknown;
  sub :ICollection;
begin
   result := 0;
   i := coll.iterator;
   while i.hasNext do begin
       obj := i.next;
       if recurse = rt_donotrecurse then
          inc(result)
       else begin
           obj.queryInterface(ICollection, sub);
           if sub <> nil then
              inc(result, count(sub, recurse))
           else
              inc(result, 1);
       end
   end
end;

function count( map  :IMap; recurse :recurse_t = rt_donotrecurse) :Integer; overload;
var
  i   :IIterator;
  obj :IUnknown;
  sub :IMap;
begin
   result := 0;
   i := map.values.iterator;
   while i.hasNext do begin
       obj := i.next;
       if recurse = rt_donotrecurse then
          inc(result)
       else begin
           obj.queryInterface(IMap, sub);
           if sub <> nil then
              inc(result, count(sub, recurse))
           else
              inc(result, 1);
       end
   end
end;


end.
