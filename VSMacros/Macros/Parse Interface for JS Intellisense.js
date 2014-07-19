// Parse Interface for JS Intellisense
/* 
    -------------------------------------------------------
    Parses
    -------------------------------------------------------
        //
        // Summary:
        //     Gets the current endpoint of the selection.
        //
        // Returns:
        //     A EnvDTE.VirtualPoint object.
        [DispId(4)]
        VirtualPoint ActivePoint { get; }
    -------------------------------------------------------
    Into
    -------------------------------------------------------
        // Gets the current endpoint of the selection.
        this.ActivePoint = ""
    -------------------------------------------------------
*/

dte.ActiveDocument.Selection.StartOfLine(0, false);
dte.ActiveDocument.Selection.LineDown(true, 2);
dte.ActiveDocument.Selection.Delete(1);
dte.ActiveDocument.Selection.CharRight(false, 2);
dte.ActiveDocument.Selection.WordRight(true, 1);
dte.ActiveDocument.Selection.Insert(" ", 1)
dte.ActiveDocument.Selection.LineDown(false, 1);
dte.ActiveDocument.Selection.CharLeft(false, 2);
dte.ActiveDocument.Selection.LineDown(true, 4);
dte.ActiveDocument.Selection.Delete(1);
dte.ActiveDocument.Selection.WordRight(true, 1);
dte.ActiveDocument.Selection.Insert("this.", 1);
dte.ActiveDocument.Selection.DeleteLeft(1);
dte.ActiveDocument.Selection.Insert(".", 1)
dte.ActiveDocument.Selection.CharRight(true, 2);
dte.ActiveDocument.Selection.WordRight(false, 1);
dte.ActiveDocument.Selection.EndOfLine(true);
dte.ActiveDocument.Selection.Insert("= \"\";", 1);
dte.ActiveDocument.Selection.NewLine(1);
dte.ActiveDocument.Selection.LineDown(false, 1);
