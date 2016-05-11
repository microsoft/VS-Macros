// Inserts a for loop to iterate over all windows

Macro.InsertText("for (var i = 1; i <= dte.Windows.Count; i++) {");
dte.ActiveDocument.Selection.NewLine();
Macro.InsertText("var doc = dte.Windows.Item(i);");
dte.ActiveDocument.Selection.NewLine(2);