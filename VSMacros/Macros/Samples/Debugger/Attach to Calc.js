// Attaches to calc.exe if it is running.

var attached = false;

for (var i = 1; i <= dte.Debugger.LocalProcesses.Count; i++) {
    var proc = dte.Debugger.LocalProcesses.Item(i);

    if (proc.Name == "calc.exe")
    {
        proc.Attach();
        attached = true;
    }
}