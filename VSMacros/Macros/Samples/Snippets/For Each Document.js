// Inserts a for loop to iterate over all documents

Macro.InsertText("for (var i = 1; i <= dte.Documents.Count; i++) {");
dte.ActiveDocument.Selection.NewLine();
Macro.InsertText("var doc = dte.Documents.Item(i);");
dte.ActiveDocument.Selection.NewLine(2);