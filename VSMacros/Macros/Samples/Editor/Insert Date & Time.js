// Inserts the date and time at the cursor position.

var date = new Date();

var day = date.getDate();
var month = date.getMonth() + 1;
var year = date.getYear();

var hours = date.getHours();
var minutes = date.getMinutes();

// Add a zero if single digit
if (day <= 9) day = "0" + day;
if (month<= 9) month = "0" + month;
if (minutes <= 9) minutes = "0" + minutes;
if (hours <= 9) hours = "0" + hours;

Macro.InsertText(month + "/" + day + "/" + year + ", " + hours + ":" + minutes);