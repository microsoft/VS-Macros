// Inserts an Undo Context and places the cursor for quick naming

Macro.InsertText("dte.UndoContext.Open(\"\");");
dte.ActiveDocument.Selection.NewLine(2);
Macro.InsertText("dte.UndoContext.Close();");
dte.ActiveDocument.Selection.LineUp(false, 2);
dte.ActiveDocument.Selection.CharLeft(false, 2);