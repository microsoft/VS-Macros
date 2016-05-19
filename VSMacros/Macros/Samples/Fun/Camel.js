/* Spawns a zombie that will eat your bad code RAWRRRRR.
 This macro is dedicated to our manager Anson Horton. */

if (dte.UndoContext.IsOpen)
    dte.UndoContext.Close();

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
dte.ActiveDocument.Selection.StartOfLine();
Macro.InsertText(camel);

dte.UndoContext.Close();