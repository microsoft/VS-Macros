/*
* Because who doesn't like camels?
*/

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
Macro.InsertText(camel);
dte.UndoContext.Close();