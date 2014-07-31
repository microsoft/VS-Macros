// Inserts the time at the cursor position.

var date = new Date();

var hours = date.getHours();
var minutes = date.getMinutes();

// Add a zero if single digit
if (minutes <= 9) minutes = "0" + minutes;
if (hours <= 9) hours = "0" + hours;

Macro.InsertText(hours + ":" + minutes);