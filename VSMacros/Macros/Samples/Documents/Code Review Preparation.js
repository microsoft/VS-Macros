/// <reference path="C:\Users\t-grawa\AppData\Local\Microsoft\VisualStudio\12.0\Extensions\sceqhjxm.rtc\Intellisense\dte.js" />
// Adds headers, remove and sorts usings in all C# files in the solution.

iterateFiles();

function iterateFiles() {
    for (var i = 1; i <= dte.Solution.Projects.Count; i++) {
        iterateProjectFiles(dte.Solution.Projects.Item(i).ProjectItems);
    }
}

function iterateProjectFiles(projectItems) {
    for (var i = 1; i <= projectItems.Count; i++) {
        var file = projectItems.Item(i);

        if (file.SubProject != null) {
            formatFile(file);
            iterateProjectFiles(file.ProjectItems);
        } else if (file.ProjectItems != null && file.ProjectItems.Count > 0) {
            formatFile(file);
            iterateProjectFiles(file.ProjectItems);
        } else {
            formatFile(file);
        }
    }
}

function formatFile(file) {
    dte.ExecuteCommand("View.SolutionExplorer");
    if (file.Name.indexOf(".cs", file.Name.length - ".cs".length) !== -1) {
        file.Open();
        file.Document.Activate();

        // Format the document
        dte.ActiveDocument.Selection.StartOfDocument(false);
        dte.ActiveDocument.Selection.EndOfLine(true);

        var text = dte.ActiveDocument.Selection.Text;
        var re = /\/\/-*/;

        // Go to start of document
        dte.ActiveDocument.Selection.StartOfDocument(false);

        dte.ExecuteCommand("Edit.RemoveAndSort");
        if (!re.test(text))
            insertHeader(dte.ActiveDocument);

        file.Document.Save();
        file.Document.Close();
    }
}

// Helper function to create a header
function insertHeader(doc) {
    var MakeDivider = function () {
        doc.Selection.Insert("//-----------------------------------------------------------------------", 1);
        doc.Selection.NewLine();
    }

    var filename = doc.Name;

    if (dte.UndoContext.IsOpen)
        dte.UndoContext.Close();

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
}