// Inserts code to modify a property using dte

Macro.InsertText("var property = dte.Properties(\"\", \"\");");
dte.ActiveDocument.Selection.NewLine();
Macro.InsertText("property.Item(\"\").Value = 0;");
dte.ActiveDocument.Selection.LineUp();
dte.ActiveDocument.Selection.CharRight(false, 3);