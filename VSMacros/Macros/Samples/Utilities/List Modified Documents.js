// Lists in the Output Window all the open documents that are currently dirty (need to be saved)

var vsWindowKindCommandWindow = "{28836128-FC2C-11D2-A433-00C04F72D18A}";
var vsWindowKindOutput = "{34E76E81-EE4A-11D0-AE2E-00A0C90FFFC3}";
var target, window;

var GetOutputWindowPane = function(name)
{
    window = dte.Windows.Item(vsWindowKindOutput);
    window.Visible = true;

    var outputWindow = window.Object;
    var outputWindowPane = outputWindow.OutputWindowPanes.Add(name);

    outputWindowPane.Activate();
    return outputWindowPane;
}

window = dte.Windows.Item(vsWindowKindCommandWindow);

if (typeof (dte.ActiveWindow) == window)
{
    target = window.Object;
}
else
{
    target = GetOutputWindowPane("Modified Documents");
    target.clear();
}

for (var i = 1; i < dte.Documents.Count; i++) {
    var doc = dte.Documents.Item(i);

    if (!doc.Saved)
    {
        target.OutputString(doc.Name + "   " + doc.FullName + "\n");
    }
}