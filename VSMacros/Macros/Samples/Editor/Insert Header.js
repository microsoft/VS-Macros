// Creates a standard copyright header

var doc = dte.ActiveDocument;

// Helper function to create a header
var MakeDivider = function () {
    doc.Selection.Insert("//-----------------------------------------------------------------------", 1);
    doc.Selection.NewLine();
}

var filename = doc.Name;

if (dte.UndoContext.IsOpen)
    dte.UndoContext.Close();

dte.UndoContext.Open("Insert Header");

// Go to start of document
doc.Selection.StartOfDocument(false);

// Start header
MakeDivider();

// Insert header content
doc.Selection.Insert("// <copyright file=\"", 1);
doc.Selection.Insert(filename, 1);
doc.Selection.Insert("\" company=\"Microsoft Corporation\">\n", 1);
doc.Selection.Insert("//     Copyright Microsoft Corporation. All rights reserved.\n", 1);
doc.Selection.Insert("// </copyright>\n", 1);

// Close header
MakeDivider();

doc.Selection.NewLine();

dte.UndoContext.Close();