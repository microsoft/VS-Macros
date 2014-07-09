// The top-level object in the Visual Studio automation object model.
var dte = new function () {

    // Gets the active document.
    this.ActiveDocument = new function () {

        // Gets the collection containing the EnvDTE.Document object.
        this.Collection;

        // Gets the top-level extensibility object.
        this.DTE;

        // Gets the Extender category ID (CATID) for the object.
        this.ExtenderCATID;

        // Gets a list of available Extenders for the object.
        this.ExtenderNames;

        // Gets the full path and name of the object's file.
        this.FullName;

        // Microsoft Internal Use Only.
        this.IndentSize;

        // Gets a GUID string indicating the kind or type of the object.
        this.Kind;

        // Microsoft Internal Use Only.
        this.Language;

        // Gets the name of the EnvDTE.Document.
        this.Name;

        // Gets the path, without file name, for the directory containing the document.
        this.Path;
        // Gets the EnvDTE.ProjectItem object associated with the EnvDTE.Document object.
        this.ProjectItem;
        // Microsoft Internal Use Only.
        this.ReadOnly;

        // Returns true if the object has not been modified since last being saved or opened.
        this.Saved;

        // Gets an object representing the current selection on the EnvDTE.Document.
        this.Selection;

        // Microsoft Internal Use Only.
        this.TabSize;

        // Microsoft Internal Use Only.
        this.Type;

        // Gets a EnvDTE.Windows collection containing the windows that display in the object.
        this.Windows;

        // Moves the focus to the current item.
        this.Activate = function () { }

        // Microsoft Internal Use Only.
        this.ClearBookmarks = function () { }

        // Closes the open document and optionally saves it, or closes and destroys the
        this.NewWindow = function () { }

        // Returns an interface or object that can be accessed at run time by name.
        this.PrintOut = function () { }

        // Re-executes the last action that was undone by the EnvDTE.Document.Undo method
        this.Redo = function () { }

        // Microsoft Internal Use Only.
        this.Undo = function () { 
    }

    // Gets an array of currently selected projects.
    this.ActiveSolutionProjects = function () {
    }

    // Returns the currently active window, or the top-most window if no others are active.
    this.ActiveWindow = function () {
    }

    // Gets the EnvDTE.AddIns collection, which contains all currently available Add-ins.
    this.AddIns = function () {
    }

    // Gets a reference to the development environment's command bars.
    this.CommandBars = function () {
    }

    // Gets a string representing the command line arguments.
    this.CommandLineArguments = function () {
    }

    // Returns the EnvDTE.Commands collection.
    this.Commands = function () {
    }

    // Gets a EnvDTE.ContextAttributes collection which allows automation clients to add new attributes to the current selected items in the Dynamic Help window and provide contextual help for the additional attributes.
    this.ContextAttributes = function () {
    }

    // Gets the debugger objects.
    this.Debugger = function () {
    }

    // Gets the display mode, either MDI or Tabbed Documents.
    this.DisplayMode = function () {
    }

    // Gets the collection of open documents in the development environment.
    this.Documents = function () {
    }

    // Gets the top-level extensibility object.
    this.DTE = function () {
    }

    // Gets a description of the edition of the environment.
    this.Edition = function () {
    }

    // Gets a reference to the EnvDTE.Events object.
    this.Events = function () {
    }

    // Microsoft Internal Use Only.
    this.FileName = function () {
    }

    // Gets the EnvDTE.Find object that represents global text find operations.
    this.Find = function () {
    }

    // Gets the full path and name of the object's file.
    this.FullName = function () {
    }

    // Gets the EnvDTE.Globals object that contains Add-in values that may be saved in the solution (.sln) file, the project file, or in the user's profile data.
    this.Globals = function () {
    }

    // Gets the EnvDTE.ItemOperations object.
    this.ItemOperations = function () {
    }

    // Gets the ID of the locale in which the development environment is running.
    this.LocaleID = function () {
    }

    // Gets a EnvDTE.Window object representing the main development environment window.
    this.MainWindow = function () {
    }

    // Gets the mode of the development environment, either debug or design.
    this.Mode = function () {
    }

    // Sets or gets the name of the EnvDTE._DTE object.
    this.Name = function () {
    }

    // Gets the EnvDTE.ObjectExtenders object.
    this.ObjectExtenders = function () {
    }

    // Gets a string with the path to the root of the Visual Studio registry settings.
    this.RegistryRoot = function () {
    }

    // Gets a collection containing the items currently selected in the environment.
    this.SelectedItems = function () {
    }

    // Gets the EnvDTE.Solution object that represents all open projects in the current instance of the environment and allows access to the build objects.
    this.Solution = function () {
    }

    // Gets a EnvDTE.SourceControl object that allows you to manipulate the source code control state of the file behind the object.
    this.SourceControl = function () {
    }

    // Gets the EnvDTE.StatusBar object, representing the status bar on the main development environment window.
    this.StatusBar = function () {
    }

    // Gets or sets whether UI should be displayed during the execution of automation code.
    this.SuppressUI = function () {
    }

    // Gets the global EnvDTE.UndoContext object.
    this.UndoContext = function () {
    }

    // Sets or gets a value indicating whether the environment was launched by a user or by automation.
    this.UserControl = function () {
    }

    // Gets the host application's version number.
    this.Version = function () {
    }

    // Gets the EnvDTE.WindowConfigurations collection, representing all available window configurations.
    this.WindowConfigurations = function () {
    }

    // Gets a EnvDTE.Windows collection containing the windows that display in the object.
    this.Windows = function () {
    }

    this.ExecuteCommand = function (CommandName, CommandArgs) {
        /// <summary>Executes the specified command.</summary>
        /// <param name="CommandName" type="String">Required. The name of the command to invoke.</param>
        /// <param name="CommandArgs" type="String">Optional. A string containing the same arguments you would supply if you were invoking the command from the Command window. If a string is supplied, it is passed to the command line as the command's first argument and is parsed to form the various arguments for the command. This is similar to how commands are invoked in the Command window.</param>
    }
}