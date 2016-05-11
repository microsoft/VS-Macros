// Increases the font size used within the editor.
// Useful to bind to a shortcut

var fontSizeIncrement = 1;
var textEditorFontsAndColors = dte.Properties("FontsAndColors", "TextEditor");

textEditorFontsAndColors.Item("FontSize").Value += fontSizeIncrement;