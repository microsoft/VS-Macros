// Inserts the time at the cursor position.

var date = new Date();

var day = date.getDate();
var month = date.getMonth();
var year = date.getYear();

dte.ActiveDocument.Selection.Text = month + "/" + day + "/" + year;