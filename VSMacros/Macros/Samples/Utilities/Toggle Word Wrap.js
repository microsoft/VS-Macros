// Toggles the word wrapping in the text editor view

var lang = ["Basic", "Plaintext", "CSharp", "HTML", "C/C++", "XML", "JavaScript"];

for (var i = 0; i < lang.length; i++) {
    var currentValue = dte.Properties("TextEditor", lang[i]).Item("WordWrap").Value;
    dte.Properties("TextEditor", lang[i]).Item("WordWrap").Value = !currentValue
}