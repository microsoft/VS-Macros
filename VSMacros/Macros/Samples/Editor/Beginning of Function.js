// BeginningOfFunction moves the caret to the beginning of the containing definition.

var textSelection = dte.ActiveDocument.Selection;

// Define Visual Studio constants
var vsCMElementFunction = 2;
var vsCMPartHeader = 4;

var codeElement = textSelection.ActivePoint.CodeElement(vsCMElementFunction);

if (codeElement != null)
{
    textSelection.MoveToPoint(codeElement.GetStartPoint(vsCMPartHeader));
    dte.ActiveDocument.Activate();
}