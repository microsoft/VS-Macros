// The top-level object in the Visual Studio automation object model.
var dte = {
    // Gets the active document.
    ActiveDocument: "",
    ActiveSolutionProjects: "",
    ActiveWindow: "",
    AddIns: "",
    ExecuteCommand: function (CommandName, CommandArgs) {
        /// <summary>Executes the specified command.</summary>
        /// <param name="CommandName" type="String">Required. The name of the command to invoke.</param>
        /// <param name="CommandArgs" type="String">Optional. A string containing the same arguments you would supply if you were invoking the command from the Command window. If a string is supplied, it is passed to the command line as the command's first argument and is parsed to form the various arguments for the command. This is similar to how commands are invoked in the Command window.</param>
    }
}

dte.ExecuteCommand()
dte.ExecuteCommand()