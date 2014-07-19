// Places the caret at the center of the screen.

var vsPaneShowTop = 1;

var textPoint = dte.ActiveDocument.Selection.ActivePoint;
textPoint.TryToShow(vsPaneShowTop);