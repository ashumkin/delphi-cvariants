{ $Id: TestDUnitGUI.dpr,v 1.14 2002/10/06 10:12:13 neuromancer Exp $ }
{: DUnit: An XTreme testing framework for Delphi programs.
   @author  The DUnit Group.
   @version $Revision: 1.14 $
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
 * and Juancarlo A�ez.
 * Portions created The Initial Developers are Copyright (C) 1999-2000.
 * Portions created by The DUnit Group are Copyright (C) 2000.
 * All rights reserved.
 *
 * Contributor(s):
 * Kent Beck <kentbeck@csi.com>
 * Erich Gamma <Erich_Gamma@oti.com>
 * Juanco A�ez <juanco@users.sourceforge.net>
 * Chris Morris <chrismo@users.sourceforge.net>
 * Jeff Moore <JeffMoore@users.sourceforge.net>
 * Kris Golko <neuromancer@users.sourceforge.net>
 * The DUnit group at SourceForge <http://dunit.sourceforge.net>
 *
 *)

{$IFDEF LINUX}
{$DEFINE DUNIT_CLX}
{$ENDIF}

{ Define DUNIT_CLX for project to use D6+ CLX }

program CVariants.Tests.GUI;
uses
  SysUtils,
  {$IFDEF DUNIT_CLX}
  QGUITestRunner,
  {$ELSE}
  GUITestRunner,
  {$ENDIF}
  CVariants.Collections.Tests in 'CVariants.Collections.Tests.pas',
  CVariants.Tests in 'CVariants.Tests.pas',
  CVariants.Utils.Tests in 'CVariants.Utils.Tests.pas';

{$R *.res}

begin
  TGUITestRunner.runRegisteredTests;
end.
