// Removes and sorts usings for all c# files in solution

// Iterate over all files
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

        dte.ExecuteCommand("Edit.RemoveAndSort");

        // -------------------

        file.Document.Save();
        file.Document.Close();
    }
}