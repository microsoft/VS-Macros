// Decreases the font size used within the editor.
// Useful to bind to a shortcut

var fontSizeIncrement = 1;
var minimumSupportedEditorSize = 3;
var textEditorFontsAndColors = dte.Properties("FontsAndColors", "TextEditor");
var fontSize = textEditorFontsAndColors.Item("FontSize");

if (fontSize.Value >= minimumSupportedEditorSize)
{
    fontSize.Value -= fontSizeIncrement;
}