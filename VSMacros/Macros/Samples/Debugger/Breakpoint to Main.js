// Sets a pending breakpoint at the function named "main". It marks the 
// breakpoint as one set by automation.

var bps = dte.Debugger.Breakpoints.Add("main");

for (var i = 1; i <= bps.Count; i++) {
    bps.Item(i).Tag = "SetByMacro";
}