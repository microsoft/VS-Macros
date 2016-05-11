// Places the caret at the center of the screen.

var vsPaneShowCentered = 0;

var textPoint = dte.ActiveDocument.Selection.ActivePoint;
textPoint.TryToShow(vsPaneShowCentered);