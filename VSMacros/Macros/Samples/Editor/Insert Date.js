// Inserts the time at the cursor position.

var date = new Date();

var day = date.getDate();
var month = date.getMonth() + 1;
var year = date.getYear();

Macro.InsertText(month + "/" + day + "/" + year);