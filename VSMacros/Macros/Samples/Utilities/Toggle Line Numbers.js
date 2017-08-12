// Toggles line numbering for common languages

var lang = ["Basic", "Plaintext", "CSharp", "HTML", "C/C++", "XML", "JavaScript"];

for (var i = 0; i < lang.length; i++) {
    try {
        var currentValue = dte.Properties("TextEditor", lang[i]).Item("ShowLineNumbers").Value;
        dte.Properties("TextEditor", lang[i]).Item("ShowLineNumbers").Value = !currentValue
    } catch(err) {
        // Swallow exception and move on
    }
}