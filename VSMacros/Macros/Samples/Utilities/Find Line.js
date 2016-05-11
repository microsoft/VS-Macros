// Launches the Find dialog, seeding pattern with current line

var textSelection = dte.ActiveDocument.Selection;
textSelection.SelectLine();
dte.ExecuteCommand("Edit.Find");
dte.Find.FindWhat = textSelection.Text;