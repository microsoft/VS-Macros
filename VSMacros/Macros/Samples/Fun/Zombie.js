/* Spawns a zombie that will eat your bad code RAWRRRRR.
 This macro is dedicated to our manager Anson Horton. */

if (dte.UndoContext.IsOpen)
    dte.UndoContext.Close();

dte.UndoContext.Open("Zombie");

var zombie =
[
'                  .....            ',
'                 C C  /            ',
'                /<   /             ',
' ___ __________/_#__=o             ',
'/(- /(\\_\\________   \\              ',
'\\ ) \\ )_      \\o     \\             ',
'/|\\ /|\\       |\'     |             ',
'              |     _|             ',
'              /o   __\\             ',
'             / \'     |             ',
'            / /      |             ',
'           /_/\\______|             ',
'          (   _(    <              ',
'           \\    \\    \\             ',
'            \\    \\    |            ',
'             \\____\\____\\           ',
'             ____\\_\\__\\_\\          ',
'           /`   /`     o\\          ',
'           |___ |_______|.. . '
].join('\n');

Macro.InsertText("/*\n*/");
dte.ExecuteCommand("Edit.LineOpenAbove");
dte.ActiveDocument.Selection.StartOfLine();
Macro.InsertText(zombie);
dte.UndoContext.Close();