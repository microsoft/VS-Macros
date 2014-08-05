/* Spawns a zombie that will eat your bad code RAWRRRRR.
 This macro is dedicated to our manager Anson Horton. */

dte.UndoContext.Open("Camel");

var camel =
[
'     //',
'   _oo\\',
'  (__/ \\ ',
'     \\  \\/ \\/ \\',
'     (         )\\',
'      \_______/  \\',
'       [[] [[]',
'       [[] [[]'
].join('\n');

Macro.InsertText("/*\n*/");
dte.ExecuteCommand("Edit.LineOpenAbove");
Macro.InsertText(camel);

dte.UndoContext.Close();