/* Spawns a zombie that will eat your bad code RAWRRRRR.
 This macro is dedicated to our manager Anson Horton. */

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
Macro.InsertText(zombie);
Macro.SurroundText("/*", "*/", zombie);
dte.UndoContext.Close();