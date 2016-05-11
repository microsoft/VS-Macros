// Closes all editor windows except the current one.

for (var i = 1; i <= dte.Documents.Count;) {
    var doc = dte.Documents.Item(i);

    // Close if not the current document
    if (dte.ActiveDocument.FullName != doc.FullName)
        doc.Close();
    else
        i++;
}