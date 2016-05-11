// Saves a copy of the current document to a name formed
// from the doc's name appended with ".bak"

var fileName = dte.ActiveDocument.FullName + ".bak";

if (typeof (dte.ActiveDocument) != undefined)
{
    var document = dte.ActiveDocument.Object();
    var startPoint = document.StartPoint.CreateEditPoint();
    var endPoint = document.EndPoint.CreateEditPoint();
    var text = startPoint.GetText(endPoint);

    var saveChangesNo = 2;

    // Create the temp document, save, then close
    dte.ItemOperations.NewFile("General\\Text File");
    dte.ActiveDocument.Object("TextDocument").Selection.Insert(text);
    dte.ActiveDocument.Save(fileName);
    dte.ActiveDocument.Close(saveChangesNo)     // Already saved with line above, don't need to save again.
}